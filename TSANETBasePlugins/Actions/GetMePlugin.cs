using Microsoft.Xrm.Sdk;
using System;

public class GetMePlugin : IPlugin
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
            // Create an instance of the CommonIntegrationPlugin
            var integrationPlugin = new CommonIntegrationPlugin(service, tracingService);

            // Retrieve the access token
            var accessToken = integrationPlugin.Login().Result;

            // Call the GetMe method
            ApiResponse response = integrationPlugin.GetMe(accessToken).Result;

            // Set the output parameters
            context.OutputParameters["IsError"] = response.IsError;
            // Return the raw JSON response to the context output parameters
            context.OutputParameters["MeResponse"] = response.Content;
        }
        catch (Exception ex)
        {
            tracingService.Trace($"Exception: {ex.Message}");
            throw new InvalidPluginExecutionException($"An error occurred in the GetMePlugin: {ex.Message}", ex);
        }
    }
}