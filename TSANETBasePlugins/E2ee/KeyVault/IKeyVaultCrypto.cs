namespace TSANETBasePlugins.E2ee
{
    /// <summary>Key Vault operations needed by the E2EE helpers.</summary>
    public interface IKeyVaultCrypto
    {
        /// <summary>ES256-signs a SHA-256 digest with the configured signing key; returns raw r||s (64 bytes).</summary>
        byte[] Sign(byte[] digest);

        /// <summary>RSA-OAEP-256 unwraps a CEK with the key identified by <paramref name="kid"/>.</summary>
        byte[] Unwrap(string kid, byte[] encryptedKey);
    }
}
