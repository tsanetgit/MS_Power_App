using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

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
            Entity caseEntity = service.Retrieve(caseReference.LogicalName, caseReference.Id, new ColumnSet("ap_name"));

            if (!caseEntity.Contains("ap_name"))
            {
                throw new InvalidPluginExecutionException("ap_name is missing in the related tsanetcase record.");
            }

            int caseId = int.Parse(caseEntity["ap_name"].ToString());

            // Initialize the common integration plugin
            CommonIntegrationPlugin commonIntegration = new CommonIntegrationPlugin(service, tracingService);

            // Login and get access token
            string accessToken = commonIntegration.Login().Result;

            // Send case note details to the API
            ApiResponse response = commonIntegration.PostCaseNote(caseId, summary, description, priority, accessToken).Result;
            if (response.IsError)
            {
                throw new InvalidPluginExecutionException(response.Content);
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
