using Microsoft.Xrm.Sdk;
using System;

public class PostCaseNotePlugin : IPlugin
{
    public void Execute(IServiceProvider serviceProvider)
    {
        ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
        IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
        IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
        IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

        try
        {
            // Retrieve case note details from input parameters
            string summary = (string)context.InputParameters["Summary"];
            string description = (string)context.InputParameters["Description"];
            string priority = (string)context.InputParameters["Priority"];
            int caseId = (int)context.InputParameters["CaseID"];

            // Initialize the common integration plugin
            CommonIntegrationPlugin commonIntegration = new CommonIntegrationPlugin(service, tracingService);

            // Login and get access token
            string accessToken = commonIntegration.Login().Result;

            // Send case note details to the API
            ApiResponse response = commonIntegration.PostCaseNote(caseId, summary, description, priority, accessToken).Result;
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
