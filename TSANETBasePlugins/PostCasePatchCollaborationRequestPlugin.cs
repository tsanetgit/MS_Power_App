using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;

public class PostCasePatchCollaborationRequestPlugin : IPlugin
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

            // Retrieve the complete entity to get all required fields
            Entity tsacase = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(
                "ap_direction", "ownerid", "ap_tsacasetoken", "ap_caseid"));

            // Check if ap_direction is out
            if (!tsacase.Contains("ap_direction") ||
                !(tsacase["ap_direction"] is OptionSetValue) ||
                ((OptionSetValue)tsacase["ap_direction"]).Value != 1)
            {
                tracingService.Trace("PatchCollaborationRequestPlugin: ap_direction is not 1, exiting");
                return;
            }

            // Check if we have the required fields
            if (!tsacase.Contains("ownerid") || !tsacase.Contains("ap_tsacasetoken") || !tsacase.Contains("ap_caseid"))
            {
                tracingService.Trace("PatchCollaborationRequestPlugin: Missing required fields (ownerid or ap_tsacasetoken or ap_caseid)");
                return;
            }

            // Get the case token
            string caseToken = tsacase["ap_tsacasetoken"].ToString();

            // Get internal case number from the related incident if exists
            string internalCaseNumber = "";
            if (tsacase.Contains("ap_caseid") && tsacase["ap_caseid"] is EntityReference incidentReference)
            {
                // Retrieve the incident to get the ticket number
                Entity incident = service.Retrieve("incident", incidentReference.Id, new ColumnSet("ticketnumber"));
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
                return; // Exit if no incident is related
            }

            // Get owner details
            EntityReference ownerRef = (EntityReference)tsacase["ownerid"];
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
                tracingService.Trace("PatchCollaborationRequestPlugin: Successfully updated collaboration request");
            }
        }
    }
}