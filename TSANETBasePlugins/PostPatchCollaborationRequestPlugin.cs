using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;

/// <summary>
/// Plugin to handle the patch collaboration request action
/// This plugin is designed to be registered on a custom action
/// </summary>
public class PostPatchCollaborationRequestPlugin : IPlugin
{
    public void Execute(IServiceProvider serviceProvider)
    {
        ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
        IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
        IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
        IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

        tracingService.Trace("PostPatchCollaborationRequestPlugin: Execution started");

        try
        {
            // Retrieve case details from input parameters
            string caseToken = (string)context.InputParameters["CaseToken"];
            string internalCaseNumber = (string)context.InputParameters["InternalCaseNumber"];
            string submitterName = (string)context.InputParameters["SubmitterName"];
            string submitterEmail = (string)context.InputParameters["SubmitterEmail"];
            string submitterPhone = (string)context.InputParameters["SubmitterPhone"];

            // Create the submitter contact details object
            SubmitterContactDetails submitterContactDetails = new SubmitterContactDetails
            {
                Name = submitterName,
                Email = submitterEmail,
                Phone = submitterPhone
            };

            // Initialize the common integration plugin
            CommonIntegrationPlugin commonIntegration = new CommonIntegrationPlugin(service, tracingService);

            // Login and get access token
            string accessToken = commonIntegration.Login().Result;

            tracingService.Trace("Sending PATCH request to update collaboration request");

            // Send the update to the API
            ApiResponse response = commonIntegration.PatchCollaborationRequest(
                caseToken,
                internalCaseNumber,
                submitterContactDetails,
                accessToken).Result;

            // Set output parameters for action response
            context.OutputParameters["IsError"] = response.IsError;
            context.OutputParameters["Response"] = response.Content;
        }
        catch (Exception ex)
        {
            tracingService.Trace($"Exception: {ex.Message}");
            throw;
        }
    }
}