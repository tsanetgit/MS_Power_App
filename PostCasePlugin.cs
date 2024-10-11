using Microsoft.Xrm.Sdk;
using System;

public class PostCasePlugin : IPlugin
{
    public void Execute(IServiceProvider serviceProvider)
    {
        ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
        IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
        IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
        IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

        try
        {
            // Retrieve case details from input parameters
            string caseDetails = (string)context.InputParameters["CaseDetails"];

            if (string.IsNullOrEmpty(caseDetails))
            {
                throw new InvalidPluginExecutionException("Case details are required.");
            }

            // Initialize the common integration plugin
            CommonIntegrationPlugin commonIntegration = new CommonIntegrationPlugin(service, tracingService);

            // Login and get access token
            string accessToken = commonIntegration.Login().Result;

            // Send case details to the API
            string response = commonIntegration.PostCase(caseDetails, accessToken).Result;

            // Return the raw JSON response to the context output parameters
            context.OutputParameters["PostCaseResponse"] = response;
        }
        catch (Exception ex)
        {
            tracingService.Trace($"Exception: {ex.Message}");
            throw;
        }
    }
}
