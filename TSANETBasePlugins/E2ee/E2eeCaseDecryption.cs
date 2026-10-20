using Microsoft.Xrm.Sdk;
using Newtonsoft.Json.Linq;
using System;

namespace TSANETBasePlugins.E2ee
{
    /// <summary>
    /// Verifies and decrypts incoming case and note payloads, including each caseNotes item.
    /// Plaintext payloads are returned unchanged. The sender is the manifest iss: when it
    /// equals own 'Iss' (ap_value) the signature is checked against the own Key Vault signing key; otherwise
    /// against the JWKS in ap_tsanetsetting 'signingKeyCounterParty' (ap_value = sender id, ap_uri = JWKS),
    /// selected by the manifest header kid.
    /// </summary>
    public static class E2eeCaseDecryption
    {
        /// <summary>Verifies a case or standalone note, including independently signed caseNotes.</summary>
        public static string Decrypt(IServiceProvider serviceProvider, string payloadJson,
            PayloadType type = PayloadType.CollaborationRequest, string expectedRef = null)
        {
            return Decrypt(payloadJson, type, expectedRef,
                (payloadType, json, reference) => DecryptPayload(serviceProvider, payloadType, json, reference));
        }

        /// <summary>Processes the response shape with a verifier that returns plaintext or throws.</summary>
        public static string Decrypt(string payloadJson, PayloadType type, string expectedRef,
            Func<PayloadType, string, string, string> decryptPayload)
        {
            if (decryptPayload == null) throw new ArgumentNullException(nameof(decryptPayload));
            JObject payload = E2eeUtil.ParseObject(payloadJson, "Response");
            string reference = type == PayloadType.Note
                ? expectedRef
                : E2eeUtil.StringOrNull(payload["submitterCaseNumber"]);
            bool changed = false;
            if (RequiresDecryption(payload))
            {
                E2eeUtil.RequireValue(reference, "Parent submitterCaseNumber");
                payload = E2eeUtil.ParseObject(decryptPayload(type, payloadJson, reference), "Decrypted response");
                changed = true;
            }

            if (type == PayloadType.CollaborationRequest)
            {
                JToken notesToken = payload["caseNotes"];
                if (notesToken != null && notesToken.Type != JTokenType.Null)
                {
                    var notes = notesToken as JArray;
                    if (notes == null) throw new InvalidPluginExecutionException("caseNotes must be an array.");
                    for (int i = 0; i < notes.Count; i++)
                    {
                        var note = notes[i] as JObject;
                        if (note == null) throw new InvalidPluginExecutionException("caseNotes items must be objects.");
                        if (!RequiresDecryption(note)) continue;
                        E2eeUtil.RequireValue(reference, "Parent submitterCaseNumber");
                        notes[i] = E2eeUtil.ParseObject(
                            decryptPayload(PayloadType.Note, E2eeUtil.Serialize(note), reference), "Decrypted note");
                        changed = true;
                    }
                }
            }
            return changed ? E2eeUtil.Serialize(payload) : payloadJson;
        }

        private static bool RequiresDecryption(JObject payload)
        {
            JToken encrypted = payload["encrypted"];
            if (encrypted != null && encrypted.Type != JTokenType.Null && encrypted.Type != JTokenType.Boolean)
            {
                throw new InvalidPluginExecutionException("encrypted must be a boolean.");
            }
            JToken manifest = payload["manifest"];
            bool hasManifest = manifest != null && manifest.Type != JTokenType.Null
                && (manifest.Type != JTokenType.String || !string.IsNullOrEmpty((string)manifest));
            bool isEncrypted = encrypted != null && encrypted.Type == JTokenType.Boolean && (bool)encrypted;
            if (isEncrypted || hasManifest)
            {
                if (manifest == null || manifest.Type != JTokenType.String || string.IsNullOrWhiteSpace((string)manifest))
                {
                    throw new InvalidPluginExecutionException("Encrypted response has no valid manifest.");
                }
                return true;
            }
            return false;
        }

        private static string DecryptPayload(IServiceProvider serviceProvider, PayloadType type, string payloadJson, string reference)
        {
            var tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            JObject probe = E2eeUtil.ParseObject(payloadJson, "Response");
            JToken manifest = probe["manifest"];
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService systemService = CommonIntegrationPlugin.CreateSystemContextService(serviceFactory);

            Entity issSetting = E2eeUtil.GetSetting(systemService, "Iss");
            if (issSetting == null)
            {
                throw new InvalidPluginExecutionException("Response is encrypted but the 'Iss' setting is missing.");
            }

            string own = E2eeUtil.RequireValue(issSetting.GetAttributeValue<string>("ap_value"), "Iss setting (ap_value)");

            E2eePayloadDecryptor decryptor = E2eeHelperFactory.CreateDecryptor(serviceProvider);
            ManifestClaims claims = decryptor.ReadClaimsUnverified(payloadJson);
            string sender = E2eeUtil.RequireValue(claims.Iss, "manifest iss");
            string senderKid = ReadManifestKid((string)manifest);

            string senderJwk;
            string counterpart;
            if (sender == own)
            {
                counterpart = E2eeUtil.RequireValue(claims.Aud, "manifest aud");
                E2eeConfig config = E2eeHelperFactory.ReadConfig(serviceProvider);
                config.ValidateForDecrypt();
                senderJwk = new KeyVaultKeyClient(serviceProvider).GetEcPublicKeyJwk(config.SigningKeyId);
            }
            else
            {
                counterpart = sender;
                Entity signingSetting = E2eeUtil.GetSetting(systemService, "signingKeyCounterParty", sender);
                if (signingSetting == null)
                {
                    throw new InvalidPluginExecutionException("No 'signingKeyCounterParty' setting for sender '" + sender + "'.");
                }
                string jwksUrl = E2eeUtil.RequireValue(signingSetting.GetAttributeValue<string>("ap_uri"), "signingKeyCounterParty (ap_uri)");
                senderJwk = FindSigningKey(E2eeCaseEncryption.FetchJwks(jwksUrl), senderKid);
            }

            DecryptResult result = decryptor.Decrypt(type, payloadJson, new DecryptionContext
            {
                SenderSigningKeyJwk = senderJwk,
                OwnCompanyId = own,
                CounterpartCompanyId = counterpart,
                ExpectedRef = reference,
                E2eeRequired = true
            });

            if (!result.IsVerified)
            {
                E2eeUtil.Trace(tracing, "Decryption failed: " + result.Status + " (jti " + claims.Jti + ")");
                throw new InvalidPluginExecutionException("Encrypted " + type + " response could not be verified/decrypted: " + result.Status);
            }
            return result.PayloadJson;
        }

        private static string ReadManifestKid(string compactJws)
        {
            string[] parts = compactJws.Split('.');
            if (parts.Length != 3)
            {
                throw new InvalidPluginExecutionException("Manifest is not a compact JWS.");
            }
            JObject header = JObject.Parse(System.Text.Encoding.UTF8.GetString(Base64UrlEncoding.Decode(parts[0])));
            return E2eeUtil.RequireValue(E2eeUtil.StringOrNull(header["kid"]), "manifest kid");
        }

        private static string FindSigningKey(JObject jwks, string kid)
        {
            var keys = jwks["keys"] as JArray;
            if (keys != null)
            {
                foreach (JToken key in keys)
                {
                    if (E2eeUtil.StringOrNull(key["kid"]) == kid && E2eeUtil.StringOrNull(key["kty"]) == "EC")
                    {
                        return key.ToString();
                    }
                }
            }
            throw new InvalidPluginExecutionException("Sender signing key '" + kid + "' not found in the counterpart JWKS.");
        }
    }
}
