using Microsoft.Xrm.Sdk;
using System;
using System.Security.Cryptography;
using System.Text;

public class HmacSha256Plugin : IPlugin
{
    public void Execute(IServiceProvider serviceProvider)
    {
        // Obtain the tracing service
        ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

        // Obtain the execution context from the service provider
        IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

        try
        {
            tracingService.Trace("HmacSha256Plugin execution started");

            // Get input parameters
            if (!context.InputParameters.Contains("Secret") || context.InputParameters["Secret"] == null)
            {
                throw new InvalidPluginExecutionException("Required parameter 'Secret' is missing.");
            }

            if (!context.InputParameters.Contains("Body") || context.InputParameters["Body"] == null)
            {
                throw new InvalidPluginExecutionException("Required parameter 'Body' is missing.");
            }

            string secret = context.InputParameters["Secret"].ToString();
            string body = context.InputParameters["Body"].ToString();

            tracingService.Trace("Input parameters retrieved successfully");

            // Compute HMAC-SHA256
            string hmacHash = ComputeHmacSha256(secret, body);

            tracingService.Trace("HMAC-SHA256 computed successfully");

            // Set output parameter
            context.OutputParameters["HashResult"] = "sha256=" + hmacHash;

            tracingService.Trace("HmacSha256Plugin execution completed successfully");
        }
        catch (Exception ex)
        {
            tracingService.Trace($"Exception: {ex.Message}");
            throw new InvalidPluginExecutionException($"An error occurred in HmacSha256Plugin: {ex.Message}", ex);
        }
    }

    private string ComputeHmacSha256(string secret, string body)
    {
        byte[] secretBytes = Encoding.UTF8.GetBytes(secret);
        byte[] bodyBytes = Encoding.UTF8.GetBytes(body);

        using (HMACSHA256 hmac = new HMACSHA256(secretBytes))
        {
            byte[] hashBytes = hmac.ComputeHash(bodyBytes);
            return BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
        }
    }
}