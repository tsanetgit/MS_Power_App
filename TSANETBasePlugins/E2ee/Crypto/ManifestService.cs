using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace TSANETBasePlugins.E2ee
{
    /// <summary>Builds, signs, parses and verifies the compact-JWS (ES256) manifest.</summary>
    internal sealed class ManifestService
    {
        private readonly IKeyVaultCrypto _keyVault;
        private readonly ITracingService _tracing;

        public ManifestService(IKeyVaultCrypto keyVault, ITracingService tracing)
        {
            _keyVault = keyVault;
            _tracing = tracing;
        }

        public string Build(E2eeConfig config, EncryptionContext ctx, string jti, long iat, IList<ManifestFieldClaim> fields)
        {
            var header = new JObject
            {
                ["alg"] = "ES256",
                ["kid"] = config.SigningKeyId,
                ["typ"] = "tsanet-manifest+jws"
            };

            var claims = new JObject
            {
                ["ver"] = E2eeUtil.Profile,
                ["iss"] = ctx.Iss,
                ["aud"] = ctx.Aud,
                ["jti"] = jti,
                ["iat"] = iat,
                ["ref"] = ctx.Ref
            };
            if (!string.IsNullOrEmpty(ctx.Cid))
            {
                claims["cid"] = ctx.Cid;
            }
            var fieldArray = new JArray();
            foreach (ManifestFieldClaim f in fields)
            {
                fieldArray.Add(new JObject { ["name"] = f.Name, ["alg"] = f.Alg, ["dig"] = f.Dig });
            }
            claims["fields"] = fieldArray;

            string signingInput = Base64UrlEncoding.Encode(Encoding.UTF8.GetBytes(E2eeUtil.Serialize(header)))
                + "." + Base64UrlEncoding.Encode(Encoding.UTF8.GetBytes(E2eeUtil.Serialize(claims)));
            byte[] digest = E2eeUtil.Sha256(Encoding.ASCII.GetBytes(signingInput));

            E2eeUtil.Trace(_tracing, "Signing manifest via Key Vault");
            byte[] signature = _keyVault.Sign(digest);
            if (signature == null || signature.Length != 64)
            {
                throw new InvalidPluginExecutionException("Key Vault returned an invalid ES256 signature (expected 64 bytes r||s).");
            }
            return signingInput + "." + Base64UrlEncoding.Encode(signature);
        }

        public static string ManifestDigest(string manifest)
        {
            return Base64UrlEncoding.Encode(E2eeUtil.Sha256(Encoding.ASCII.GetBytes(manifest)));
        }

        /// <summary>Parses the claims WITHOUT verifying the signature. Throws on malformed input.</summary>
        public ManifestClaims ParseUnverified(string manifest)
        {
            string[] parts = (manifest ?? string.Empty).Split('.');
            if (parts.Length != 3)
            {
                throw new InvalidPluginExecutionException("Manifest is not a compact JWS.");
            }
            ManifestClaims claims = TryParseClaims(parts[1]);
            if (claims == null)
            {
                throw new InvalidPluginExecutionException("Manifest claims are malformed.");
            }
            return claims;
        }

        /// <summary>Verifies the ES256 signature locally with the supplied sender JWK.</summary>
        public DecryptStatus VerifySignature(string manifest, string senderJwk, out ManifestClaims claims)
        {
            claims = null;

            JObject jwk = E2eeUtil.ParseObject(senderJwk, "SenderSigningKeyJwk");
            string jwkKid = E2eeUtil.StringOrNull(jwk["kid"]);
            if (string.IsNullOrEmpty(jwkKid))
            {
                throw new InvalidPluginExecutionException("SenderSigningKeyJwk must contain 'kid'.");
            }

            string[] parts = manifest.Split('.');
            if (parts.Length != 3)
            {
                return DecryptStatus.SignatureInvalid;
            }

            byte[] signature;
            JObject header;
            try
            {
                header = JObject.Parse(Encoding.UTF8.GetString(Base64UrlEncoding.Decode(parts[0])));
                signature = Base64UrlEncoding.Decode(parts[2]);
            }
            catch (Exception ex) when (ex is JsonException || ex is FormatException)
            {
                return DecryptStatus.SignatureInvalid;
            }

            if (E2eeUtil.StringOrNull(header["alg"]) != "ES256"
                || E2eeUtil.StringOrNull(header["kid"]) != jwkKid
                || signature.Length != 64)
            {
                return DecryptStatus.SignatureInvalid;
            }

            byte[] digest = E2eeUtil.Sha256(Encoding.ASCII.GetBytes(parts[0] + "." + parts[1]));
            bool valid;
            try
            {
                using (ECDsa key = JwkHelper.EcPublicKeyFromJwk(senderJwk))
                {
                    valid = key.VerifyHash(digest, signature);
                }
            }
            catch (InvalidOperationException ex)
            {
                throw new InvalidPluginExecutionException("SenderSigningKeyJwk is not a valid EC P-256 public JWK: " + ex.Message, ex);
            }
            catch (CryptographicException)
            {
                valid = false;
            }

            if (!valid)
            {
                return DecryptStatus.SignatureInvalid;
            }

            claims = TryParseClaims(parts[1]);
            return claims == null ? DecryptStatus.BindingMismatch : DecryptStatus.Ok;
        }

        public static DecryptStatus CheckBindings(ManifestClaims claims, DecryptionContext ctx)
        {
            if (claims.Ver != E2eeUtil.Profile)
            {
                return DecryptStatus.BindingMismatch;
            }

            bool forward = claims.Iss == ctx.OwnCompanyId && claims.Aud == ctx.CounterpartCompanyId;
            bool reverse = claims.Iss == ctx.CounterpartCompanyId && claims.Aud == ctx.OwnCompanyId;
            if (!forward && !reverse)
            {
                return DecryptStatus.BindingMismatch;
            }

            if (claims.Ref != ctx.ExpectedRef)
            {
                return DecryptStatus.BindingMismatch;
            }

            if (!string.IsNullOrEmpty(claims.Cid) && !string.IsNullOrEmpty(ctx.ExpectedCid) && claims.Cid != ctx.ExpectedCid)
            {
                return DecryptStatus.BindingMismatch;
            }
            return DecryptStatus.Ok;
        }

        private static ManifestClaims TryParseClaims(string payloadSegment)
        {
            try
            {
                JObject obj = E2eeUtil.ParseObject(Encoding.UTF8.GetString(Base64UrlEncoding.Decode(payloadSegment)), "Manifest claims");
                var claims = new ManifestClaims
                {
                    Ver = E2eeUtil.StringOrNull(obj["ver"]),
                    Iss = E2eeUtil.StringOrNull(obj["iss"]),
                    Aud = E2eeUtil.StringOrNull(obj["aud"]),
                    Jti = E2eeUtil.StringOrNull(obj["jti"]),
                    Ref = E2eeUtil.StringOrNull(obj["ref"]),
                    Cid = E2eeUtil.StringOrNull(obj["cid"])
                };
                if (obj["iat"] == null || obj["iat"].Type != JTokenType.Integer)
                {
                    return null;
                }
                claims.Iat = (long)obj["iat"];

                var fields = obj["fields"] as JArray;
                if (fields == null || claims.Jti == null || claims.Ver == null || claims.Iss == null || claims.Aud == null || claims.Ref == null)
                {
                    return null;
                }
                foreach (JToken f in fields)
                {
                    string name = E2eeUtil.StringOrNull(f["name"]);
                    string alg = E2eeUtil.StringOrNull(f["alg"]);
                    string dig = E2eeUtil.StringOrNull(f["dig"]);
                    if (name == null || alg == null || dig == null)
                    {
                        return null;
                    }
                    claims.Fields.Add(new ManifestFieldClaim { Name = name, Alg = alg, Dig = dig });
                }
                return claims;
            }
            catch (Exception ex) when (ex is InvalidPluginExecutionException || ex is FormatException || ex is JsonException)
            {
                return null;
            }
        }
    }
}
