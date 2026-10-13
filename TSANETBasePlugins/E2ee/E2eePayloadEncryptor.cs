using Microsoft.Xrm.Sdk;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;

namespace TSANETBasePlugins.E2ee
{
    /// <summary>
    /// Field-level encryption plus one signed manifest per payload (profile tsanet-e2ee-1).
    /// Create one instance per Execute call; holds no request data.
    /// </summary>
    public class E2eePayloadEncryptor
    {
        private readonly IKeyVaultCrypto _keyVault;
        private readonly E2eeConfig _config;
        private readonly ITracingService _tracing;

        public E2eePayloadEncryptor(IKeyVaultCrypto keyVault, E2eeConfig config, ITracingService tracing)
        {
            if (keyVault == null) throw new ArgumentNullException(nameof(keyVault));
            if (config == null) throw new ArgumentNullException(nameof(config));
            _keyVault = keyVault;
            _config = config;
            _tracing = tracing;
        }

        public EncryptResult Encrypt(PayloadType type, string payloadJson, EncryptionContext ctx)
        {
            if (ctx == null) throw new InvalidPluginExecutionException("EncryptionContext is required.");
            _config.ValidateForEncrypt();
            E2eeUtil.RequireValue(ctx.Iss, "Iss");
            E2eeUtil.RequireValue(ctx.Aud, "Aud");
            E2eeUtil.RequireValue(ctx.Ref, "Ref");

            try
            {
                JObject payload = E2eeUtil.ParseObject(payloadJson, "Payload");
                E2eeUtil.Trace(_tracing, "Encrypt: payload parsed (" + type + ")");

                using (RsaRecipient recipient = RsaRecipient.FromJwk(ctx.RecipientEncryptionKeyJwk, "RecipientEncryptionKeyJwk"))
                using (RsaRecipient own = RsaRecipient.FromJwk(ctx.OwnEncryptionKeyJwk, "OwnEncryptionKeyJwk"))
                {
                    var recipients = new List<RsaRecipient> { recipient, own };
                    string jti = Guid.NewGuid().ToString();

                    List<ProtectedField> fields = PayloadFieldMap.Enumerate(payload, type, PayloadDirection.Outgoing);
                    E2eeUtil.Trace(_tracing, "Encrypt: " + fields.Count + " protected field(s)");

                    var claims = new List<ManifestFieldClaim>();
                    foreach (ProtectedField field in fields)
                    {
                        string jwe = JweFieldCrypto.Encrypt(field.Value, field.Id, jti, recipients);
                        field.SetValue(jwe);
                        claims.Add(new ManifestFieldClaim { Name = field.Id, Alg = "S256", Dig = E2eeUtil.Digest(jwe) });
                    }

                    long iat = (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds;
                    string manifest = new ManifestService(_keyVault, _tracing).Build(_config, ctx, jti, iat, claims);

                    payload["manifest"] = manifest;
                    payload["encrypted"] = true;
                    E2eeUtil.Trace(_tracing, "Encrypt: manifest signed, jti=" + jti);

                    return new EncryptResult
                    {
                        PayloadJson = E2eeUtil.Serialize(payload),
                        Jti = jti,
                        Manifest = manifest,
                        ManifestDigest = ManifestService.ManifestDigest(manifest)
                    };
                }
            }
            catch (InvalidPluginExecutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                E2eeUtil.Trace(_tracing, "Encrypt failed: " + ex.GetType().Name);
                throw new InvalidPluginExecutionException("E2EE encryption failed: " + ex.Message, ex);
            }
        }
    }
}
