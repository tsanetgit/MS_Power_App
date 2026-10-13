using Microsoft.Xrm.Sdk;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;

namespace TSANETBasePlugins.E2ee
{
    internal sealed class RsaRecipient : IDisposable
    {
        public string Kid { get; private set; }
        public RSA Key { get; private set; }

        private RsaRecipient() { }

        public static RsaRecipient FromJwk(string jwkJson, string what)
        {
            JObject jwk = E2eeUtil.ParseObject(jwkJson, what);
            string kid = E2eeUtil.StringOrNull(jwk["kid"]);
            if (string.IsNullOrEmpty(kid))
            {
                throw new InvalidPluginExecutionException(what + " must contain 'kid'.");
            }
            try
            {
                return new RsaRecipient { Kid = kid, Key = JwkHelper.RsaPublicKeyFromJwk(jwkJson) };
            }
            catch (Exception ex) when (ex is InvalidOperationException || ex is FormatException || ex is CryptographicException)
            {
                throw new InvalidPluginExecutionException(what + " is not a valid RSA public JWK: " + ex.Message, ex);
            }
        }

        public void Dispose()
        {
            if (Key != null)
            {
                Key.Dispose();
            }
        }
    }

    /// <summary>Per-field JWE (general JSON serialization, A256GCM, RSA-OAEP-256). No Dataverse dependencies.</summary>
    internal static class JweFieldCrypto
    {
        private const int CekSize = 32;
        private const int IvSize = 12;
        private const int TagSize = 16;

        public static string Encrypt(string plaintext, string fieldId, string jti, IList<RsaRecipient> recipients)
        {
            var header = new JObject
            {
                ["enc"] = "A256GCM",
                ["tsa_ver"] = E2eeUtil.Profile,
                ["tsa_jti"] = jti,
                ["tsa_fld"] = fieldId
            };
            string protectedHeader = Base64UrlEncoding.Encode(Encoding.UTF8.GetBytes(E2eeUtil.Serialize(header)));
            byte[] aad = Encoding.ASCII.GetBytes(protectedHeader);

            byte[] cek = new byte[CekSize];
            byte[] iv = new byte[IvSize];
            try
            {
                using (var rng = RandomNumberGenerator.Create())
                {
                    rng.GetBytes(cek);
                    rng.GetBytes(iv);
                }

                byte[][] sealedData = Jose.AesGcm.Encrypt(cek, iv, aad, Encoding.UTF8.GetBytes(plaintext));
                byte[] ciphertext = sealedData[0];
                byte[] tag = sealedData[1];
                if (tag.Length != TagSize)
                {
                    throw new CryptographicException("Unexpected GCM tag length.");
                }

                var recipientArray = new JArray();
                foreach (RsaRecipient recipient in recipients)
                {
                    byte[] wrapped = recipient.Key.Encrypt(cek, RSAEncryptionPadding.OaepSHA256);
                    recipientArray.Add(new JObject
                    {
                        ["header"] = new JObject { ["alg"] = "RSA-OAEP-256", ["kid"] = recipient.Kid },
                        ["encrypted_key"] = Base64UrlEncoding.Encode(wrapped)
                    });
                }

                var jwe = new JObject
                {
                    ["protected"] = protectedHeader,
                    ["recipients"] = recipientArray,
                    ["iv"] = Base64UrlEncoding.Encode(iv),
                    ["ciphertext"] = Base64UrlEncoding.Encode(ciphertext),
                    ["tag"] = Base64UrlEncoding.Encode(tag)
                };
                return E2eeUtil.Serialize(jwe);
            }
            finally
            {
                Array.Clear(cek, 0, cek.Length);
            }
        }

        /// <summary>
        /// Decrypts one field. Returns Ok and the plaintext, or a failure status with null plaintext.
        /// Key Vault faults propagate as exceptions.
        /// </summary>
        public static DecryptStatus Decrypt(string jweString, string expectedJti, string expectedFieldId,
            string ownKeyPrefix, string vaultPrefix, IKeyVaultCrypto keyVault, out string plaintext)
        {
            plaintext = null;

            string protectedHeader, kid;
            byte[] iv, ciphertext, tag, encryptedKey = null;
            JObject jwe;
            try
            {
                jwe = JObject.Parse(jweString);
            }
            catch (Newtonsoft.Json.JsonException)
            {
                return DecryptStatus.AeadFailure;
            }

            protectedHeader = E2eeUtil.StringOrNull(jwe["protected"]);
            string ivText = E2eeUtil.StringOrNull(jwe["iv"]);
            string ctText = E2eeUtil.StringOrNull(jwe["ciphertext"]);
            string tagText = E2eeUtil.StringOrNull(jwe["tag"]);
            var recipients = jwe["recipients"] as JArray;
            if (protectedHeader == null || ivText == null || ctText == null || tagText == null || recipients == null)
            {
                return DecryptStatus.AeadFailure;
            }

            string encryptedKeyText = null;
            kid = null;
            foreach (JToken recipient in recipients)
            {
                string candidate = E2eeUtil.StringOrNull(recipient["header"] == null ? null : recipient["header"]["kid"]);
                if (candidate != null
                    && candidate.StartsWith(vaultPrefix, StringComparison.OrdinalIgnoreCase)
                    && candidate.StartsWith(ownKeyPrefix, StringComparison.Ordinal))
                {
                    kid = candidate;
                    encryptedKeyText = E2eeUtil.StringOrNull(recipient["encrypted_key"]);
                    break;
                }
            }
            if (kid == null)
            {
                return DecryptStatus.NoMatchingRecipient;
            }

            try
            {
                iv = Base64UrlEncoding.Decode(ivText);
                ciphertext = Base64UrlEncoding.Decode(ctText);
                tag = Base64UrlEncoding.Decode(tagText);
                if (encryptedKeyText == null)
                {
                    return DecryptStatus.AeadFailure;
                }
                encryptedKey = Base64UrlEncoding.Decode(encryptedKeyText);
            }
            catch (FormatException)
            {
                return DecryptStatus.AeadFailure;
            }

            if (iv.Length != IvSize || tag.Length != TagSize)
            {
                return DecryptStatus.AeadFailure;
            }

            byte[] aad = Encoding.ASCII.GetBytes(protectedHeader);
            byte[] cek = keyVault.Unwrap(kid, encryptedKey);
            byte[] plainBytes;
            try
            {
                if (cek == null || cek.Length != CekSize)
                {
                    return DecryptStatus.AeadFailure;
                }
                try
                {
                    plainBytes = Jose.AesGcm.Decrypt(cek, iv, aad, ciphertext, tag);
                }
                catch (Exception ex) when (ex is CryptographicException || ex is Jose.JoseException)
                {
                    return DecryptStatus.AeadFailure;
                }
            }
            finally
            {
                if (cek != null)
                {
                    Array.Clear(cek, 0, cek.Length);
                }
            }

            JObject header;
            try
            {
                header = JObject.Parse(Encoding.UTF8.GetString(Base64UrlEncoding.Decode(protectedHeader)));
            }
            catch (Exception ex) when (ex is Newtonsoft.Json.JsonException || ex is FormatException)
            {
                return DecryptStatus.HeaderBindingMismatch;
            }

            if (E2eeUtil.StringOrNull(header["enc"]) != "A256GCM"
                || E2eeUtil.StringOrNull(header["tsa_ver"]) != E2eeUtil.Profile
                || E2eeUtil.StringOrNull(header["tsa_jti"]) != expectedJti
                || E2eeUtil.StringOrNull(header["tsa_fld"]) != expectedFieldId)
            {
                return DecryptStatus.HeaderBindingMismatch;
            }

            plaintext = Encoding.UTF8.GetString(plainBytes);
            return DecryptStatus.Ok;
        }
    }
}
