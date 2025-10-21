using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;

public class PostCaseNoteOnCreatePlugin : IPlugin
{
    public void Execute(IServiceProvider serviceProvider)
    {
        ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
        IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
        IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
        IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

        if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
        {
            Entity entity = (Entity)context.InputParameters["Target"];

            string noteId = entity.Contains("ap_tsanotecode") ? entity["ap_tsanotecode"].ToString() : string.Empty;

            if (!string.IsNullOrEmpty(noteId))
            {
                return;
            }

            // Retrieve case note details from the entity fields
            string summary = entity.Contains("ap_name") ? entity["ap_name"].ToString() : string.Empty;
            string description = entity.Contains("ap_description") ? entity["ap_description"].ToString() : string.Empty;
            string priority = entity.Contains("ap_priority") ? GetPriority(entity.GetAttributeValue<OptionSetValue>("ap_priority").Value) : "LOW";

            // Retrieve the related tsanetcase record using ap_tsanetcaseid
            if (!entity.Contains("ap_tsanetcaseid"))
            {
                throw new InvalidPluginExecutionException("ap_tsanetcaseid is missing.");
            }

            EntityReference caseReference = (EntityReference)entity["ap_tsanetcaseid"];
            Entity caseEntity = service.Retrieve(caseReference.LogicalName, caseReference.Id, new ColumnSet("ap_name", "ap_tsacasetoken"));

            if (!caseEntity.Contains("ap_tsacasetoken"))
            {
                throw new InvalidPluginExecutionException("ap_tsacasetoken is missing in the related tsanetcase record.");
            }

            string caseToken = caseEntity["ap_tsacasetoken"].ToString();

            // Retrieve user details
            Entity user = service.Retrieve("systemuser", context.UserId, new ColumnSet("firstname", "lastname", "address1_telephone1", "internalemailaddress", "address1_city", "domainname"));
            string userFirstName = user.GetAttributeValue<string>("firstname");
            string userLastName = user.GetAttributeValue<string>("lastname");
            string userName = user.GetAttributeValue<string>("domainname");
            string userPhone = user.GetAttributeValue<string>("address1_telephone1");
            string userEmail = user.GetAttributeValue<string>("internalemailaddress");
            string userCity = user.GetAttributeValue<string>("address1_city");

            // Initialize the common integration plugin
            CommonIntegrationPlugin commonIntegration = new CommonIntegrationPlugin(service, tracingService);

            // Login and get access token
            string accessToken = commonIntegration.Login().Result;

            // Create SubmittedBy object
            SubmittedBy submittedBy = new SubmittedBy
            {
                Username = userName,
                FirstName = userFirstName,
                LastName = userLastName,
                Phone = userPhone,
                Email = userEmail,
                City = userCity
            };

            // Send case note details to the API
            ApiResponse response = commonIntegration.PostCaseNote(caseToken, summary, description, priority, submittedBy, accessToken).Result;
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
                var caseNote = JsonConvert.DeserializeObject<CaseNote>(response.Content);

                if (caseNote != null)
                {
                    Entity noteUpdate = new Entity(entity.LogicalName, entity.Id);
                    // Update the ap_tsanotecode field with the caseNoteId
                    noteUpdate["ap_tsanotecode"] = caseNote.Id.ToString();
                    noteUpdate["ap_creatorname"] = caseNote.CreatorName;
                    noteUpdate["ap_creatoremail"] = caseNote.CreatorEmail;
                    noteUpdate["ap_companyname"] = caseNote.CompanyName;
                    service.Update(noteUpdate);
                }
            }
        }
    }

    private string GetPriority(int priorityValue)
    {
        switch (priorityValue)
        {
            case 1:
                return "HIGH";
            case 2:
                return "MEDIUM";
            case 3:
                return "LOW";
            default:
                return "LOW";
        }
    }
}
