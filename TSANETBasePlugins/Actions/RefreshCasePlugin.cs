using Microsoft.Xrm.Sdk;
using System;

public class RefreshCasePlugin : IPlugin
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
            EntityReference targetRef = null;

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
            {
                targetRef = (EntityReference)context.InputParameters["Target"];
                tracingService.Trace($"Executing RefreshCasePlugin for entity: {targetRef.LogicalName} with ID: {targetRef.Id}");
            }
            else
            {
                tracingService.Trace("No Target entity found in context");
            }

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

                if (!response.IsError && targetRef != null)
                {
                    tracingService.Trace("Case data retrieved successfully, processing response");

                    try
                    {
                        // Use the UpdateCasePlugin to process the response
                        var commonCasePlugin = new CommonCasePlugin();
                        commonCasePlugin.ProcessCaseResponse(service, tracingService, response.Content, targetRef.Id);

                        tracingService.Trace("Case data processed successfully");
                        context.OutputParameters["Success"] = true;
                        context.OutputParameters["Message"] = "Case refreshed successfully";
                    }
                    catch (Exception ex)
                    {
                        tracingService.Trace($"Error processing case data: {ex.Message}");
                        context.OutputParameters["Success"] = false;
                        context.OutputParameters["Message"] = $"Error processing case data: {ex.Message}";
                    }
                }

                // Return the raw JSON response to the context output parameters
                context.OutputParameters["CaseResponse"] = response.Content;
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
