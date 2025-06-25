using Microsoft.Xrm.Sdk;
using System;

public class GetCasePlugin : IPlugin
{
    public void Execute(IServiceProvider serviceProvider)
    {
        // Obtain the tracing service
        ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

        // Obtain the execution context from the service provider.
        IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

        // Obtain the organization service reference.
        IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
        IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

        try
        {
            // Extract the internal case number from the input parameters
            if (context.InputParameters.Contains("CaseToken") && context.InputParameters["CaseToken"] is string caseToken)
            {
                // Create an instance of the CommonIntegrationPlugin
                var integrationPlugin = new CommonIntegrationPlugin(service, tracingService);

                // Retrieve the access token
                var accessToken = integrationPlugin.Login().Result;

                // Call the GetCaseU method
                ApiResponse response = integrationPlugin.GetCase(caseToken, accessToken).Result;

                // Set the output parameter
                context.OutputParameters["IsError"] = response.IsError;
                // Return the raw JSON response to the context output parameters
                context.OutputParameters["GetCaseResponse"] = response.Content;
            }
            else
            {
                throw new InvalidPluginExecutionException("Token parameter is missing.");
            }
        }
        catch (Exception ex)
        {
            tracingService.Trace($"Exception: {ex.Message}");
            throw new InvalidPluginExecutionException($"An error occurred in the GetCasePlugin: {ex.Message}", ex);
        }
    }
}
