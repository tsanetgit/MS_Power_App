using System;

namespace TSANETBasePlugins.E2ee
{
    /// <summary>
    /// HTTP implementation backed by the existing <see cref="KeyVaultKeyClient"/>
    /// (managed identity token, scope https://vault.azure.net/.default).
    /// </summary>
    internal class KeyVaultCryptoHttp : IKeyVaultCrypto
    {
        private readonly KeyVaultKeyClient _client;
        private readonly E2eeConfig _config;

        public KeyVaultCryptoHttp(IServiceProvider serviceProvider, E2eeConfig config)
        {
            _client = new KeyVaultKeyClient(serviceProvider);
            _config = config;
        }

        public byte[] Sign(byte[] digest)
        {
            return _client.Sign(_config.SigningKeyId, digest);
        }

        public byte[] Unwrap(string kid, byte[] encryptedKey)
        {
            return _client.UnwrapKey(kid, encryptedKey);
        }
    }
}
