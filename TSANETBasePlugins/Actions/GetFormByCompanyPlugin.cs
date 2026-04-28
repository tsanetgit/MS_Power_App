using Microsoft.Xrm.Sdk;
using System;

public class GetFormByCompanyPlugin : IPlugin
{
    public void Execute(IServiceProvider serviceProvider)
    {
        ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
        IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
        IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
        IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

        try
        {
            // Retrieve the company ID from input parameters as an integer
            int companyId = (int)context.InputParameters["CompanyId"];
            if (companyId <= 0)
            {
                throw new InvalidPluginExecutionException("Valid Company ID is required.");
            }

            // Initialize the common integration plugin
            CommonIntegrationPlugin commonIntegration = new CommonIntegrationPlugin(service, tracingService);

            // Login and get access token
            string accessToken = commonIntegration.Login().Result;

            // Get form details by company ID
            ApiResponse response = commonIntegration.GetFormByCompany(companyId, accessToken).Result;
            context.OutputParameters["IsError"] = response.IsError;

            if (response.IsError)
            {
                // Return error
                context.OutputParameters["ErrorMessage"] = response.Content;
            }
            else
            {
                // Return the raw JSON response to the context output parameters
                context.OutputParameters["FormDetails"] = response.Content;
            }

        }
        catch (Exception ex)
        {
            tracingService.Trace($"Exception: {ex.Message}");
            throw;
        }
    }
}
