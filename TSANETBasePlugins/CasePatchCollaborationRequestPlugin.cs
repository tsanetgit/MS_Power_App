using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;

public class CasePatchCollaborationRequestPlugin : IPlugin
{
    public void Execute(IServiceProvider serviceProvider)
    {
        ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
        IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
        IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
        IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

        tracingService.Trace("PatchCollaborationRequestPlugin: Execution started");

        if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
        {
            Entity entity = (Entity)context.InputParameters["Target"];

            // Get the pre-image if available
            Entity preImage = null;
            if (context.PreEntityImages.Contains("PreImage"))
            {
                preImage = context.PreEntityImages["PreImage"];
            }

            // Define a method to get attribute from either target or pre-image
            T GetAttributeValue<T>(string attributeName) where T : class
            {
                if (entity.Contains(attributeName) && entity[attributeName] is T)
                {
                    return (T)entity[attributeName];
                }
                else if (preImage != null && preImage.Contains(attributeName) && preImage[attributeName] is T)
                {
                    return (T)preImage[attributeName];
                }
                return null;
            }

            /// Check if ap_direction is out (value = 1)
            OptionSetValue directionValue = GetAttributeValue<OptionSetValue>("ap_direction");
            if (directionValue == null || directionValue.Value != 1)
            {
                tracingService.Trace("PatchCollaborationRequestPlugin: ap_direction is not 1, exiting");
                return;
            }

            // Check for required fields
            EntityReference ownerRef = GetAttributeValue<EntityReference>("ownerid");
            string caseToken = null;
            if (entity.Contains("ap_tsacasetoken"))
                caseToken = entity["ap_tsacasetoken"].ToString();
            else if (preImage != null && preImage.Contains("ap_tsacasetoken"))
                caseToken = preImage["ap_tsacasetoken"].ToString();

            EntityReference caseIdRef = GetAttributeValue<EntityReference>("ap_caseid");

            if (ownerRef == null || string.IsNullOrEmpty(caseToken) || caseIdRef == null)
            {
                tracingService.Trace("PatchCollaborationRequestPlugin: Missing required fields (ownerid or ap_tsacasetoken or ap_caseid)");
                return;
            }

            // Get internal case number from the related incident if exists
            string internalCaseNumber = "";
            if (caseIdRef != null)
            {
                // Retrieve the incident to get the ticket number
                Entity incident = service.Retrieve("incident", caseIdRef.Id, new ColumnSet("ticketnumber"));
                if (incident.Contains("ticketnumber"))
                {
                    internalCaseNumber = incident["ticketnumber"].ToString();
                    tracingService.Trace($"Retrieved internal case number from incident: {internalCaseNumber}");
                }
                else
                {
                    tracingService.Trace("No ticketnumber found on the incident");
                    return; // Exit if no ticket number is found
                }
            }
            else
            {
                tracingService.Trace("No related incident found");
            }

            // Get owner details
            Entity ownerEntity;

            // Owner could be a user or team
            if (ownerRef.LogicalName == "systemuser")
            {
                ownerEntity = service.Retrieve("systemuser", ownerRef.Id, new ColumnSet(
                    "firstname", "lastname", "fullname", "internalemailaddress", "address1_telephone1"));
            }
            else if (ownerRef.LogicalName == "team")
            {
                tracingService.Trace("PatchCollaborationRequestPlugin: Owner is a team, which is not currently supported");
                return;
            }
            else
            {
                tracingService.Trace($"PatchCollaborationRequestPlugin: Unexpected owner type {ownerRef.LogicalName}");
                return;
            }

            // Create the submitter contact details object
            SubmitterContactDetails submitterContactDetails = new SubmitterContactDetails
            {
                Name = ownerEntity.Contains("fullname") ?
                       ownerEntity["fullname"].ToString() :
                       $"{ownerEntity.GetAttributeValue<string>("firstname")} {ownerEntity.GetAttributeValue<string>("lastname")}",
                Email = ownerEntity.GetAttributeValue<string>("internalemailaddress"),
                Phone = ownerEntity.GetAttributeValue<string>("address1_telephone1")
            };

            // Initialize the common integration plugin
            CommonIntegrationPlugin commonIntegration = new CommonIntegrationPlugin(service, tracingService);

            // Login and get access token
            string accessToken = commonIntegration.Login().Result;

            tracingService.Trace("PatchCollaborationRequestPlugin: Sending PATCH request to update collaboration request");

            // Send the update to the API
            ApiResponse response = commonIntegration.PatchCollaborationRequest(
                caseToken,
                internalCaseNumber,
                submitterContactDetails,
                accessToken).Result;

            // Process response
            if (response.IsError)
            {
                tracingService.Trace($"PatchCollaborationRequestPlugin: Error response received - {response.Content}");
                var errorResponse = JsonConvert.DeserializeObject<Dictionary<string, string>>(response.Content);
                if (errorResponse != null && errorResponse.ContainsKey("message"))
                {
                    throw new InvalidPluginExecutionException(errorResponse["message"]);
                }
                else
                {
                    throw new InvalidPluginExecutionException("An unknown error occurred when updating the collaboration request.");
                }
            }
            else
            {
                entity["ap_formjson"] = response.Content;
                tracingService.Trace("PatchCollaborationRequestPlugin: Successfully updated case form");

                CommonCasePlugin commonCasePlugin = new CommonCasePlugin();
                commonCasePlugin.ProcessCaseResponse(service, tracingService, response.Content, entity.Id);

                tracingService.Trace("PatchCollaborationRequestPlugin: Successfully updated collaboration request");
            }
        }
    }
}