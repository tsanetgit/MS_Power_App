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
            // Retrieve the company ID from input parameters
            string companyId = (string)context.InputParameters["CompanyId"];
            if (string.IsNullOrEmpty(companyId))
            {
                throw new InvalidPluginExecutionException("Company ID is required.");
            }

            // Initialize the common integration plugin
            CommonIntegrationPlugin commonIntegration = new CommonIntegrationPlugin(service);

            // Login and get access token
            string accessToken = commonIntegration.Login().Result;

            // Get form details by company ID
            string formJson = commonIntegration.GetFormByCompany(companyId, accessToken).Result;

            // Return the raw JSON response to the context output parameters
            context.OutputParameters["FormDetails"] = formJson;
        }
        catch (Exception ex)
        {
            tracingService.Trace($"Exception: {ex.Message}");
            throw;
        }
    }
}
