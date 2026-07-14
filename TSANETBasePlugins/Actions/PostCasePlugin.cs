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
        // Retrieve case details from input parameters
        string caseDetails = (string)context.InputParameters["CaseDetails"];
        string caseID = (string)context.InputParameters["CaseID"];

        if (string.IsNullOrEmpty(caseDetails))
        {
            throw new InvalidPluginExecutionException("Case details are required.");
        }

        // Initialize the common integration plugin
        CommonIntegrationPlugin commonIntegration = new CommonIntegrationPlugin(serviceProvider);

        // Login and get access token
        string accessToken = commonIntegration.Login().Result;

        // Send case details to the API
        ApiResponse response = commonIntegration.PostCase(caseDetails, accessToken).Result;
        context.OutputParameters["IsError"] = response.IsError;

        tracingService.Trace("CaseID is: " + caseID);

        if (!response.IsError && !string.IsNullOrEmpty(caseID))
        {
            tracingService.Trace("Case data retrieved successfully, processing response");

            // Use the UpdateCasePlugin to process the response
            var commonCasePlugin = new CommonCasePlugin();
            commonCasePlugin.ProcessCaseResponse(service, tracingService, response.Content, new Guid(caseID));

            tracingService.Trace("Case data processed successfully");
        }
        // Return the raw JSON response to the context output parameters
        context.OutputParameters["PostCaseResponse"] = response.Content;


    }
}
