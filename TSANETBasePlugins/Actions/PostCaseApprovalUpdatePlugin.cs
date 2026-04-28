using Microsoft.Xrm.Sdk;
using System;

public class PostCaseApprovalUpdatePlugin : IPlugin
{
    public void Execute(IServiceProvider serviceProvider)
    {
        ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
        IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
        IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
        IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

        try
        {
            // Retrieve case approval update details from input parameters
            string caseNumber = (string)context.InputParameters["CaseNumber"];
            string engineerName = (string)context.InputParameters["EngineerName"];
            string engineerPhone = (string)context.InputParameters["EngineerPhone"];
            string engineerEmail = (string)context.InputParameters["EngineerEmail"];
            string nextSteps = (string)context.InputParameters["NextSteps"];
            int caseId = (int)context.InputParameters["CaseID"];

            // Initialize the common integration plugin
            CommonIntegrationPlugin commonIntegration = new CommonIntegrationPlugin(service, tracingService);

            // Login and get access token
            string accessToken = commonIntegration.Login().Result;

            // Send case approval update details to the API
            ApiResponse response = commonIntegration.UpdateCaseApproval(caseId, caseNumber, engineerName, engineerPhone, engineerEmail, nextSteps, accessToken).Result;
            context.OutputParameters["IsError"] = response.IsError;
            // Return the raw JSON response to the context output parameters
            context.OutputParameters["Response"] = response.Content;
        }
        catch (Exception ex)
        {
            tracingService.Trace($"Exception: {ex.Message}");
            throw;
        }
    }
}
