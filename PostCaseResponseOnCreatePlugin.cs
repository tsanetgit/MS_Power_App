using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;

public class PostCaseResponseOnCreatePlugin : IPlugin
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
                    Entity tsanetcase = service.Retrieve(tsanetcaseRef.LogicalName, tsanetcaseRef.Id, new ColumnSet("ap_submittercasenumber", "ap_name", "ap_tsacasetoken"));

                    // Retrieve case approval details from the related tsanetcase entity
                    string caseNumber = entity.GetAttributeValue<string>("ap_internalcasenumber");
                    string caseToken = tsanetcase["ap_tsacasetoken"].ToString();

                    // Retrieve nextSteps from the current entity
                    string description = entity.GetAttributeValue<string>("ap_description");
                    int type = entity.GetAttributeValue<OptionSetValue>("ap_type").Value;

                    // Update statecode and statuscode based on type
                    UpdateStateAndStatus(service, tsanetcase, type);

                    string respId = entity.Contains("ap_tsaresponsecode") ? entity["ap_tsaresponsecode"].ToString() : string.Empty;

                    if (!string.IsNullOrEmpty(respId))
                    {
                        return;
                    }

                    // Retrieve current user details
                    //Entity user = service.Retrieve("systemuser", context.UserId, new ColumnSet("fullname", "address1_telephone1", "internalemailaddress"));
                    string engineerName = entity.GetAttributeValue<string>("ap_engineername");
                    string engineerPhone = entity.GetAttributeValue<string>("ap_engineerphone");
                    string engineerEmail = entity.GetAttributeValue<string>("ap_engineeremail");

                    // Initialize the common integration plugin
                    CommonIntegrationPlugin commonIntegration = new CommonIntegrationPlugin(service, tracingService);

                    // Login and get access token
                    string accessToken = commonIntegration.Login().Result;
                    ApiResponse response = new ApiResponse();
                    // Send case request details to the API
                    // approval
                    if (type == 1)
                    {
                        response = commonIntegration.PostCaseApproval(caseToken, caseNumber, engineerName, engineerPhone, engineerEmail, description, accessToken).Result;
                    }
                    // reject
                    else if (type == 0)
                    {
                        response = commonIntegration.PostCaseReject(caseToken, engineerName, engineerPhone, engineerEmail, description, accessToken).Result;
                    }
                    // request information
                    else if (type == 2)
                    {
                        response = commonIntegration.PostCaseRequestInformation(caseToken, engineerName, engineerPhone, engineerEmail, description, accessToken).Result;
                    }
                    //information response
                    else if (type == 3)
                    {
                        response = commonIntegration.PostCaseInformationResponse(caseToken, engineerName, engineerPhone, engineerEmail, description, accessToken).Result;
                    }
                    //close
                    else if (type == 4)
                    {
                        response = commonIntegration.PostCaseClose(caseToken, accessToken).Result;
                    } else if (type == 5)
                    {
                        response = commonIntegration.PatchCaseApproval(caseToken, caseNumber, engineerName, engineerPhone, engineerEmail, description, accessToken).Result;
                    }
                    //Process response
                    if (response.IsError)
                    {
                        var errorResponse = JsonConvert.DeserializeObject<Dictionary<string, string>>(response.Content);
                        if (errorResponse != null && errorResponse.ContainsKey("message"))
                        {
                            throw new InvalidPluginExecutionException(errorResponse["message"]);
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException("An unknown error occurred.");
                        }
                    }
                    else
                    {
                        var caseResponse = JsonConvert.DeserializeObject<Case>(response.Content);
                        // Get the response with the highest id
                        var lastCaseResp = caseResponse.CaseResponses.OrderByDescending(caseResp => caseResp.Id).FirstOrDefault();

                        if (lastCaseResp != null)
                        {
                            // Update the code field with the last id
                            entity["ap_tsaresponsecode"] = type == 5 ? Guid.NewGuid().ToString() + ":" + lastCaseResp.Id.ToString() : lastCaseResp.Id.ToString();
                            service.Update(entity);

                            tsanetcase["ap_submittercasenumber"] = caseNumber;
                            service.Update(tsanetcase);
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

    private void UpdateStateAndStatus(IOrganizationService service, Entity entity, int type)
    {
        int statecode = 0;
        int statuscode = 1;

        switch (type)
        {
            case 0:
                statecode = 1;
                statuscode = 120950002;
                break;
            case 1:
                statecode = 0;
                statuscode = 120950003;
                break;
            case 2:
                statecode = 0;
                statuscode = 120950001;
                break;
            case 3:
                statecode = 0;
                statuscode = 1;
                break;
            case 4:
                statecode = 1;
                statuscode = 2;
                break;
            case 5:
                statecode = 0;
                statuscode = 120950003;
                break;
        }

        entity["statecode"] = new OptionSetValue(statecode);
        entity["statuscode"] = new OptionSetValue(statuscode);
        service.Update(entity);
    }
}
