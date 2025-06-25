using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Net.Http;
using System.Threading.Tasks;

public class GetCompanyPlugin : IPlugin
{
    public void Execute(IServiceProvider serviceProvider)
    {
        ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
        IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
        IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
        IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

        try
        {
            // Retrieve the company name from input parameters
            string companyName = (string)context.InputParameters["CompanyName"];
            if (string.IsNullOrEmpty(companyName))
            {
                throw new InvalidPluginExecutionException("Company name is required.");
            }

            // Initialize the common integration plugin
            CommonIntegrationPlugin commonIntegration = new CommonIntegrationPlugin(service, tracingService);

            // Login and get access token
            string accessToken = commonIntegration.Login().Result;

            // Get company details by name
            string companyDetailsJson = commonIntegration.GetCompanyByName(companyName, accessToken).Result;

            // Return the raw JSON response to the context output parameters
            context.OutputParameters["CompanyDetails"] = companyDetailsJson;
        }
        catch (Exception ex)
        {
            tracingService.Trace($"Exception: {ex.Message}");
            throw;
        }
    }
}
