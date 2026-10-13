using Microsoft.Xrm.Sdk;
using Newtonsoft.Json.Linq;
using System;

namespace TSANETBasePlugins.E2ee
{
    /// <summary>
    /// Verifies and decrypts an incoming case payload when it carries an E2EE manifest.
    /// Payloads without a manifest are returned unchanged. The sender is the manifest iss: when it
    /// equals own 'Iss' (ap_value) the signature is checked against the own Key Vault signing key; otherwise
    /// against the JWKS in ap_tsanetsetting 'signingKeyCounterParty' (ap_value = sender id, ap_uri = JWKS),
    /// selected by the manifest header kid.
    /// </summary>
    public static class E2eeCaseDecryption
    {
        /// <summary>Returns plaintext payload JSON; throws if a manifest is present but verification or decryption fails.</summary>
        public static string Decrypt(IServiceProvider serviceProvider, string payloadJson)
        {
            var tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            JObject probe;
            try
            {
                probe = JObject.Parse(payloadJson);
            }
            catch (Exception)
            {
                return payloadJson;
            }
            JToken encrypted = probe["encrypted"];
            if (encrypted == null || encrypted.Type != JTokenType.Boolean || !(bool)encrypted)
            {
                return payloadJson;
            }
            JToken manifest = probe["manifest"];
            if (manifest == null || manifest.Type != JTokenType.String || string.IsNullOrEmpty((string)manifest))
            {
                return payloadJson;
            }

            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService systemService = CommonIntegrationPlugin.CreateSystemContextService(serviceFactory);

            Entity issSetting = E2eeUtil.GetSetting(systemService, "Iss");            if (issSetting == null)
            {
                throw new InvalidPluginExecutionException("Response is encrypted but the 'Iss' setting is missing.");
            }

            string own = E2eeUtil.RequireValue(issSetting.GetAttributeValue<string>("ap_value"), "Iss setting (ap_value)");
            string reference = E2eeUtil.RequireValue(E2eeUtil.StringOrNull(probe["submitterCaseNumber"]), "submitterCaseNumber");

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

            DecryptResult result = decryptor.Decrypt(PayloadType.CollaborationRequest, payloadJson, new DecryptionContext
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
                throw new InvalidPluginExecutionException("Encrypted case response could not be verified/decrypted: " + result.Status);
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
