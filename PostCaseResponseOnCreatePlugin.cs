using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Linq;

public class PostCaseApprovalOnCreatePlugin : IPlugin
{
    public void Execute(IServiceProvider serviceProvider)
    {
        ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
        IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
        IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
        IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

        try
        {
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity entity = (Entity)context.InputParameters["Target"];

                // Retrieve related tsanetcase details
                if (entity.Contains("ap_tsanetcaseid") && entity["ap_tsanetcaseid"] is EntityReference)
                {
                    EntityReference tsanetcaseRef = (EntityReference)entity["ap_tsanetcaseid"];
                    Entity tsanetcase = service.Retrieve(tsanetcaseRef.LogicalName, tsanetcaseRef.Id, new ColumnSet("ap_submittercasenumber", "ap_name"));

                    // Retrieve case approval details from the related tsanetcase entity
                    string caseNumber = tsanetcase.GetAttributeValue<string>("ap_submittercasenumber");
                    int caseId = int.Parse(tsanetcase["ap_name"].ToString());

                    // Retrieve nextSteps from the current entity
                    string nextSteps = entity.GetAttributeValue<string>("ap_description");

                    // Retrieve current user details
                    Entity user = service.Retrieve("systemuser", context.UserId, new ColumnSet("fullname", "address1_telephone1", "internalemailaddress"));
                    string engineerName = user.GetAttributeValue<string>("fullname");
                    string engineerPhone = user.GetAttributeValue<string>("address1_telephone1");
                    string engineerEmail = user.GetAttributeValue<string>("internalemailaddress");

                    // Initialize the common integration plugin
                    CommonIntegrationPlugin commonIntegration = new CommonIntegrationPlugin(service, tracingService);

                    // Login and get access token
                    string accessToken = commonIntegration.Login().Result;

                    // Send case approval details to the API
                    ApiResponse response = commonIntegration.PostCaseApproval(caseId, caseNumber, engineerName, engineerPhone, engineerEmail, nextSteps, accessToken).Result;
                    if (response.IsError)
                    {
                        throw new InvalidPluginExecutionException(response.Content);
                    }
                    else
                    {
                        var caseResponse = JsonConvert.DeserializeObject<Case>(response.Content);
                        // Get the response with the highest id
                        var lastCaseResp = caseResponse.CaseResponses.OrderByDescending(caseResp => caseResp.Id).FirstOrDefault();

                        if (lastCaseResp != null)
                        {
                            // Update the code field with the last id
                            entity["ap_tsaresponsecode"] = lastCaseResp.Id.ToString();
                            service.Update(entity);
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            tracingService.Trace($"Exception: {ex.Message}");
            throw;
        }
    }
}
