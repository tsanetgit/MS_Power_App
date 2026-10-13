using Microsoft.Xrm.Sdk;
using System;

namespace TSANETBasePlugins.E2ee
{
    /// <summary>Key Vault configuration (environment variables ap_SigningKeyId and ap_UnwrapKeyId; the vault URI is derived from the key URIs).</summary>
    public class E2eeConfig
    {
        public string KeyVaultUri { get; set; }

        /// <summary>Returns {scheme}://{host} of a Key Vault key URI, or null when it is not one.</summary>
        internal static string VaultFromKeyId(string keyId)
        {
            if (string.IsNullOrWhiteSpace(keyId))
            {
                return null;
            }
            string trimmed = keyId.Trim();
            int idx = trimmed.IndexOf("/keys/", StringComparison.OrdinalIgnoreCase);
            return idx > 0 ? trimmed.Substring(0, idx) : null;
        }

        /// <summary>Key-version URI of the ES256 signing key: {vault}/keys/{name}/{version}.</summary>
        public string SigningKeyId { get; set; }

        /// <summary>URI of the own RSA encryption key (versioned or versionless).</summary>
        public string EncryptionKeyId { get; set; }

        internal string VaultPrefix
        {
            get { return KeyVaultUri.Trim().TrimEnd('/') + "/"; }
        }

        internal void ValidateForEncrypt()
        {
            Require(SigningKeyId, "ap_SigningKeyId");
            Require(KeyVaultUri, "ap_SigningKeyId (Key Vault URI)");
            string[] parts = KeySegments(SigningKeyId, "ap_SigningKeyId");            if (parts.Length != 3)
            {
                throw new InvalidPluginExecutionException("ap_SigningKeyId must be a key-version URI ({vault}/keys/{name}/{version}).");
            }
        }

        internal void ValidateForDecrypt()
        {
            Require(EncryptionKeyId, "ap_UnwrapKeyId");
            Require(KeyVaultUri, "ap_UnwrapKeyId (Key Vault URI)");
            OwnKeyPrefix();
        }

        /// <summary>Versionless own key URI with trailing slash: {vault}/keys/{name}/.</summary>
        internal string OwnKeyPrefix()
        {
            string[] parts = KeySegments(EncryptionKeyId, "ap_UnwrapKeyId");            if (parts.Length < 2 || parts.Length > 3)
            {
                throw new InvalidPluginExecutionException("ap_UnwrapKeyId must be {vault}/keys/{name}[/{version}].");            }
            return VaultPrefix + "keys/" + parts[1] + "/";
        }

        private string[] KeySegments(string keyId, string name)
        {
            string trimmed = keyId.Trim().TrimEnd('/');
            if (!trimmed.StartsWith(VaultPrefix, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidPluginExecutionException(name + " is not under the configured Key Vault.");            }
            string[] parts = trimmed.Substring(VaultPrefix.Length).Split('/');
            if (parts.Length < 2 || parts[0] != "keys")
            {
                throw new InvalidPluginExecutionException(name + " is not a Key Vault key URI.");
            }
            return parts;
        }

        private static void Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new InvalidPluginExecutionException("Configuration value " + name + " is missing.");
            }
        }
    }
}
