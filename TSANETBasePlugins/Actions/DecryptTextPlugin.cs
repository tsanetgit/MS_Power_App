using Microsoft.Xrm.Sdk;
using System;

/// <summary>
/// Implements the unbound custom action "tsanet_DecryptText": decrypts a JWE (A256GCM /
/// RSA-OAEP-256) and verifies the inner JWS (ES256) against the sender's public key, per
/// the TSANet Connect E2E encryption proposal. Manual/testing invocation only.
/// </summary>
public class DecryptTextPlugin : IPlugin
{
    public void Execute(IServiceProvider serviceProvider)
    {
        ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
        IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

        try
        {
            if (!context.InputParameters.Contains("EncryptedText") || !(context.InputParameters["EncryptedText"] is string encryptedText) || string.IsNullOrEmpty(encryptedText))
            {
                throw new InvalidPluginExecutionException("EncryptedText is required.");
            }

            if (!context.InputParameters.Contains("SenderPublicKeyJwk") || !(context.InputParameters["SenderPublicKeyJwk"] is string senderPublicKeyJwk) || string.IsNullOrEmpty(senderPublicKeyJwk))
            {
                throw new InvalidPluginExecutionException("SenderPublicKeyJwk is required.");
            }

            var joseDecryptHelper = new JoseDecryptHelper(serviceProvider);
            (string plainText, bool verified) = joseDecryptHelper.Decrypt(encryptedText, senderPublicKeyJwk);

            // Never surface unverified plaintext — quarantine per the E2E encryption
            context.OutputParameters["PlainText"] = verified ? plainText : string.Empty;
            context.OutputParameters["VerificationStatus"] = verified;
        }
        catch (Exception ex)
        {
            tracingService.Trace($"Exception: {ex.Message}");
            throw;
        }
    }
}
