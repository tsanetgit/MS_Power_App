using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;

namespace TSANETBasePlugins.E2ee
{
    /// <summary>
    /// Encrypts an outgoing collaboration request when (and only when) an 'encryptedRecipient'
    /// record exists in ap_tsanetsetting.
    /// <list type="bullet">
    /// <item>Iss: ap_tsanetsetting 'Iss' (ap_value)</item>
    /// <item>Aud: ap_tsanetsetting 'encryptedRecipient' (ap_value)</item>
    /// <item>Recipient key: JWKS at ap_tsanetsetting 'encryptedRecipient' (ap_uri)</item>
    /// <item>Own key (self-wrap): public key of Key Vault key ap_UnwrapKeyId</item>
    /// <item>Ref: caseDetails.internalCaseNumber</item>
    /// </list>
    /// </summary>
    public static class E2eeCaseEncryption
    {
        private const string RecipientSettingName = "encryptedRecipient";
        private const string IssSettingName = "Iss";

        /// <summary>Returns the encrypted payload, or <paramref name="caseDetailsJson"/> unchanged when no recipient is configured.</summary>
        public static string PrepareForPost(IServiceProvider serviceProvider, string caseDetailsJson)
        {
            var tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService systemService = CommonIntegrationPlugin.CreateSystemContextService(serviceFactory);

            Entity recipient = E2eeUtil.GetSetting(systemService, RecipientSettingName);
            if (recipient == null)
            {
                E2eeUtil.Trace(tracing, "No 'encryptedRecipient' setting; posting case unencrypted");
                return caseDetailsJson;
            }

            string aud = recipient.GetAttributeValue<string>("ap_value");
            string jwksUrl = recipient.GetAttributeValue<string>("ap_uri");
            Entity issSetting = E2eeUtil.GetSetting(systemService, IssSettingName);
            string iss = issSetting == null ? null : issSetting.GetAttributeValue<string>("ap_value");
            E2eeUtil.RequireValue(aud, "encryptedRecipient (ap_value)");
            E2eeUtil.RequireValue(jwksUrl, "encryptedRecipient (ap_uri)");
            E2eeUtil.RequireValue(iss, "Iss setting (ap_value)");

            JObject payload = E2eeUtil.ParseObject(caseDetailsJson, "CaseDetails");
            string reference = E2eeUtil.StringOrNull(payload["internalCaseNumber"]);
            E2eeUtil.RequireValue(reference, "CaseDetails.internalCaseNumber");

            E2eeConfig config = E2eeHelperFactory.ReadConfig(serviceProvider);
            config.ValidateForEncrypt();
            config.ValidateForDecrypt();

            var ctx = new EncryptionContext
            {
                RecipientEncryptionKeyJwk = FetchRecipientJwk(jwksUrl, tracing),
                OwnEncryptionKeyJwk = new KeyVaultKeyClient(serviceProvider).GetPublicKeyJwk(config.EncryptionKeyId),
                Iss = iss,
                Aud = aud,
                Ref = reference
            };

            E2eeUtil.Trace(tracing, "Encrypting case for recipient " + aud);
            E2eePayloadEncryptor encryptor = E2eeHelperFactory.CreateEncryptor(serviceProvider);
            return encryptor.Encrypt(PayloadType.CollaborationRequest, caseDetailsJson, ctx).PayloadJson;
        }

        private static string FetchRecipientJwk(string jwksUrl, ITracingService tracing)        {
            JObject jwks = FetchJwks(jwksUrl);
            var keys = jwks["keys"] as JArray;
            if (keys != null)
            {
                foreach (JToken token in keys)
                {
                    string use = E2eeUtil.StringOrNull(token["use"]);
                    string alg = E2eeUtil.StringOrNull(token["alg"]);
                    if (E2eeUtil.StringOrNull(token["kty"]) == "RSA"
                        && !string.IsNullOrEmpty(E2eeUtil.StringOrNull(token["kid"]))
                        && (use == null || use == "enc")
                        && (alg == null || alg == "RSA-OAEP-256"))
                    {
                        E2eeUtil.Trace(tracing, "Recipient encryption key selected: " + E2eeUtil.StringOrNull(token["kid"]));
                        return token.ToString();
                    }
                }
            }
            throw new InvalidPluginExecutionException("The recipient JWKS contains no usable RSA encryption key (kty RSA, kid, use enc, RSA-OAEP-256).");
        }

        internal static JObject FetchJwks(string jwksUrl)
        {
            Uri uri;
            if (!Uri.TryCreate(jwksUrl, UriKind.Absolute, out uri) || uri.Scheme != Uri.UriSchemeHttps)
            {
                throw new InvalidPluginExecutionException("The recipient JWKS URL must be an absolute https URL.");
            }

            string body;
            try
            {
                using (var client = new HttpClient { Timeout = TimeSpan.FromSeconds(30) })
                {
                    HttpResponseMessage response = client.GetAsync(uri).GetAwaiter().GetResult();
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new InvalidPluginExecutionException("Fetching the recipient JWKS failed with status " + (int)response.StatusCode + ".");
                    }
                    body = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                }
            }
            catch (InvalidPluginExecutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Fetching the recipient JWKS failed: " + ex.Message, ex);
            }

            return E2eeUtil.ParseObject(body, "Recipient JWKS");
        }
    }
}
