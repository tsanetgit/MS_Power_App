using Microsoft.Xrm.Sdk;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TSANETBasePlugins.E2ee
{
    /// <summary>
    /// Verifies the manifest and decrypts protected fields. Create one instance per Execute call.
    /// Does not check freshness or replay and performs no Dataverse I/O.
    /// </summary>
    public class E2eePayloadDecryptor
    {
        private readonly IKeyVaultCrypto _keyVault;
        private readonly E2eeConfig _config;
        private readonly ITracingService _tracing;

        public E2eePayloadDecryptor(IKeyVaultCrypto keyVault, E2eeConfig config, ITracingService tracing)
        {
            if (keyVault == null) throw new ArgumentNullException(nameof(keyVault));
            if (config == null) throw new ArgumentNullException(nameof(config));
            _keyVault = keyVault;
            _config = config;
            _tracing = tracing;
        }

        /// <summary>
        /// Reads the manifest claims WITHOUT verifying the signature, bindings or fingerprints.
        /// The result must NEVER be trusted on its own: use it only to select keys or look up
        /// context before calling <see cref="Decrypt"/>, which performs the real verification.
        /// Throws if the manifest is absent or malformed.
        /// </summary>
        public ManifestClaims ReadClaimsUnverified(string payloadJson)
        {
            JObject payload = E2eeUtil.ParseObject(payloadJson, "Payload");
            string manifest = E2eeUtil.StringOrNull(payload["manifest"]);
            if (string.IsNullOrEmpty(manifest))
            {
                throw new InvalidPluginExecutionException("Payload has no manifest.");
            }
            return new ManifestService(_keyVault, _tracing).ParseUnverified(manifest);
        }

        public DecryptResult Decrypt(PayloadType type, string payloadJson, DecryptionContext ctx)
        {
            if (ctx == null) throw new InvalidPluginExecutionException("DecryptionContext is required.");
            _config.ValidateForDecrypt();

            try
            {
                JObject payload = E2eeUtil.ParseObject(payloadJson, "Payload");

                JToken manifestToken = payload["manifest"];
                bool noManifest = manifestToken == null || manifestToken.Type == JTokenType.Null
                    || (manifestToken.Type == JTokenType.String && string.IsNullOrEmpty((string)manifestToken));
                if (noManifest)
                {
                    E2eeUtil.Trace(_tracing, "Decrypt: no manifest; required=" + ctx.E2eeRequired);
                    return ctx.E2eeRequired
                        ? Fail(DecryptStatus.ManifestMissing)
                        : new DecryptResult { IsVerified = true, Status = DecryptStatus.NotEncrypted, PayloadJson = payloadJson };
                }

                E2eeUtil.RequireValue(ctx.OwnCompanyId, "OwnCompanyId");
                E2eeUtil.RequireValue(ctx.CounterpartCompanyId, "CounterpartCompanyId");
                E2eeUtil.RequireValue(ctx.ExpectedRef, "ExpectedRef");

                if (manifestToken.Type != JTokenType.String)
                {
                    return Fail(DecryptStatus.SignatureInvalid);
                }
                string manifest = (string)manifestToken;

                var manifests = new ManifestService(_keyVault, _tracing);
                ManifestClaims claims;
                DecryptStatus status = manifests.VerifySignature(manifest, ctx.SenderSigningKeyJwk, out claims);
                E2eeUtil.Trace(_tracing, "Decrypt: signature check = " + status);
                if (status == DecryptStatus.SignatureInvalid)
                {
                    return Fail(status);
                }

                string manifestDigest = ManifestService.ManifestDigest(manifest);
                if (status != DecryptStatus.Ok)
                {
                    return Fail(status);
                }

                status = ManifestService.CheckBindings(claims, ctx);
                E2eeUtil.Trace(_tracing, "Decrypt: bindings = " + status);
                if (status != DecryptStatus.Ok)
                {
                    return Fail(status, claims, manifestDigest);
                }

                List<ProtectedField> fields = PayloadFieldMap.Enumerate(payload, type, PayloadDirection.Incoming);
                var payloadNames = new HashSet<string>(fields.Select(f => f.Id));
                var manifestNames = new HashSet<string>(claims.Fields.Select(f => f.Name));
                if (manifestNames.Count != claims.Fields.Count || !payloadNames.SetEquals(manifestNames))
                {
                    return Fail(DecryptStatus.FieldSetMismatch, claims, manifestDigest);
                }
                E2eeUtil.Trace(_tracing, "Decrypt: field set matches (" + fields.Count + ")");

                var expected = claims.Fields.ToDictionary(f => f.Name);
                foreach (ProtectedField field in fields)
                {
                    ManifestFieldClaim claim = expected[field.Id];
                    if (claim.Alg != "S256" || !E2eeUtil.FixedTimeEquals(E2eeUtil.Digest(field.Value), claim.Dig))
                    {
                        return Fail(DecryptStatus.FingerprintMismatch, claims, manifestDigest);
                    }
                }
                E2eeUtil.Trace(_tracing, "Decrypt: fingerprints match");

                string ownPrefix = _config.OwnKeyPrefix();
                var plaintexts = new Dictionary<string, string>();
                foreach (ProtectedField field in fields)
                {
                    string plaintext;
                    DecryptStatus fieldStatus = JweFieldCrypto.Decrypt(field.Value, claims.Jti, field.Id,
                        ownPrefix, _config.VaultPrefix, _keyVault, out plaintext);
                    if (fieldStatus != DecryptStatus.Ok)
                    {
                        E2eeUtil.Trace(_tracing, "Decrypt: field '" + field.Id + "' failed = " + fieldStatus);
                        return Fail(fieldStatus, claims, manifestDigest);
                    }
                    plaintexts[field.Id] = plaintext;
                }

                JObject result = (JObject)payload.DeepClone();
                foreach (ProtectedField field in PayloadFieldMap.Enumerate(result, type, PayloadDirection.Incoming))
                {
                    field.SetValue(plaintexts[field.Id]);
                }
                E2eeUtil.Trace(_tracing, "Decrypt: verified, jti=" + claims.Jti);

                return new DecryptResult
                {
                    IsVerified = true,
                    Status = DecryptStatus.Ok,
                    PayloadJson = E2eeUtil.Serialize(result),
                    Jti = claims.Jti,
                    Iat = claims.Iat,
                    ManifestDigest = manifestDigest
                };
            }
            catch (InvalidPluginExecutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                E2eeUtil.Trace(_tracing, "Decrypt failed: " + ex.GetType().Name);
                throw new InvalidPluginExecutionException("E2EE decryption failed: " + ex.Message, ex);
            }
        }

        private static DecryptResult Fail(DecryptStatus status)
        {
            return new DecryptResult { IsVerified = false, Status = status, PayloadJson = null };
        }

        private static DecryptResult Fail(DecryptStatus status, ManifestClaims claims, string manifestDigest)
        {
            return new DecryptResult
            {
                IsVerified = false,
                Status = status,
                PayloadJson = null,
                Jti = claims.Jti,
                Iat = claims.Iat,
                ManifestDigest = manifestDigest
            };
        }
    }
}
