using Microsoft.Xrm.Sdk;
using System;

/// <summary>
/// Implements the unbound custom action "tsanet_EncryptText": signs (ES256) and then
/// encrypts (JWE, A256GCM / RSA-OAEP-256) plain text for a recipient, per the TSANet
/// Connect E2E encryption proposal. Manual/testing invocation only.
/// </summary>
public class EncryptTextPlugin : IPlugin
{
    public void Execute(IServiceProvider serviceProvider)
    {
        ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
        IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

        try
        {
            if (!context.InputParameters.Contains("PlainText") || !(context.InputParameters["PlainText"] is string plainText) || string.IsNullOrEmpty(plainText))
            {
                throw new InvalidPluginExecutionException("PlainText is required.");
            }

            if (!context.InputParameters.Contains("RecipientPublicKeyJwk") || !(context.InputParameters["RecipientPublicKeyJwk"] is string recipientPublicKeyJwk) || string.IsNullOrEmpty(recipientPublicKeyJwk))
            {
                throw new InvalidPluginExecutionException("RecipientPublicKeyJwk is required.");
            }

            var joseEncryptHelper = new JoseEncryptHelper(serviceProvider);
            string encryptedText = joseEncryptHelper.Encrypt(plainText, recipientPublicKeyJwk);

            context.OutputParameters["EncryptedText"] = encryptedText;
        }
        catch (Exception ex)
        {
            Exception traced = ex is AggregateException aggregate ? aggregate.Flatten() : ex;
            tracingService.Trace($"Exception: {traced}");
            throw;
        }
    }
}
