using Microsoft.Xrm.Sdk;
using System;

namespace TSANETBasePlugins.E2ee
{
    /// <summary>Wires the existing infrastructure (system-context env-var reads, managed-identity Key Vault client).</summary>
    public static class E2eeHelperFactory
    {
        public static E2eePayloadEncryptor CreateEncryptor(IServiceProvider serviceProvider)
        {
            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            E2eeConfig config = ReadConfig(serviceProvider);
            config.ValidateForEncrypt();
            return new E2eePayloadEncryptor(new KeyVaultCryptoHttp(serviceProvider, config), config, tracing);
        }

        public static E2eePayloadDecryptor CreateDecryptor(IServiceProvider serviceProvider)
        {
            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            E2eeConfig config = ReadConfig(serviceProvider);
            config.ValidateForDecrypt();
            return new E2eePayloadDecryptor(new KeyVaultCryptoHttp(serviceProvider, config), config, tracing);
        }

        internal static E2eeConfig ReadConfig(IServiceProvider serviceProvider)
        {
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService systemService = CommonIntegrationPlugin.CreateSystemContextService(serviceFactory);
            string signingKeyId = CommonIntegrationPlugin.GetEnvVariable(systemService, "ap_SigningKeyId");
            string unwrapKeyId = CommonIntegrationPlugin.GetEnvVariable(systemService, "ap_UnwrapKeyId");
            return new E2eeConfig
            {
                KeyVaultUri = E2eeConfig.VaultFromKeyId(signingKeyId) ?? E2eeConfig.VaultFromKeyId(unwrapKeyId),
                SigningKeyId = signingKeyId,
                EncryptionKeyId = unwrapKeyId
            };
        }
    }
}
