using Microsoft.Xrm.Sdk;
using System;
using System.Globalization;
using TSANETBasePlugins.E2ee;

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
        string recipientCompanyId = (string)context.InputParameters["RecipientCompanyID"];

        if (string.IsNullOrEmpty(caseDetails))
        {
            throw new InvalidPluginExecutionException("Case details are required.");
        }

        caseDetails = E2eeCaseEncryption.PrepareForPost(serviceProvider, PayloadType.CollaborationRequest, caseDetails, null, recipientCompanyId);
        // Initialize the common integration plugin
        CommonIntegrationPlugin commonIntegration = new CommonIntegrationPlugin(serviceProvider);

        // Login and get
        string accessToken = commonIntegration.Login().Result;

        // Send case details to the API
        ApiResponse response = commonIntegration.PostCase(caseDetails, accessToken).Result;
        var caseResponseJson = response.Content;

        context.OutputParameters["IsError"] = response.IsError;

        tracingService.Trace("CaseID is: " + caseID);

        if (!response.IsError && !string.IsNullOrEmpty(caseID))
        {
            tracingService.Trace("Case data retrieved successfully, processing response");

            // Use the UpdateCasePlugin to process the response
            var commonCasePlugin = new CommonCasePlugin();
            caseResponseJson = commonCasePlugin.ProcessCaseResponse(service, tracingService, caseResponseJson, new Guid(caseID), commonIntegration, accessToken, serviceProvider);

            tracingService.Trace("Case data processed successfully");
        }
        // Return the raw JSON response to the context output parameters
        context.OutputParameters["PostCaseResponse"] = caseResponseJson;


    }
}
