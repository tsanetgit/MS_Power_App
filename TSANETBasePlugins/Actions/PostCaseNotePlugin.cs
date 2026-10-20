using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using TSANETBasePlugins.E2ee;

public class PostCaseNotePlugin : IPlugin
{
    public void Execute(IServiceProvider serviceProvider)
    {
        ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
        IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
        IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
        IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

        try
        {
            // Retrieve case note details from input parameters
            string summary = (string)context.InputParameters["Summary"];
            string description = (string)context.InputParameters["Description"];
            string priority = (string)context.InputParameters["Priority"];
            int caseId = (int)context.InputParameters["CaseID"];
            string caseToken = (string)context.InputParameters["CaseToken"];

            // Create SubmittedBy object
            SubmittedBy submittedBy = new SubmittedBy
            {
            };

            var parentQuery = new QueryExpression("ap_tsanetcase")
            {
                ColumnSet = new ColumnSet("ap_submittercasenumber"),
                TopCount = 2
            };
            parentQuery.Criteria.AddCondition("ap_tsacasetoken", ConditionOperator.Equal, caseToken);
            EntityCollection parents = service.RetrieveMultiple(parentQuery);
            if (parents.Entities.Count != 1)
            {
                throw new InvalidPluginExecutionException("The note must resolve to exactly one parent collaboration request.");
            }
            string parentReference = parents.Entities[0].GetAttributeValue<string>("ap_submittercasenumber");
            string noteDetails = JsonConvert.SerializeObject(new { summary, description, priority, submittedBy });
            noteDetails = E2eeCaseEncryption.PrepareForPost(serviceProvider, PayloadType.Note, noteDetails, parentReference);

            // Initialize the common integration plugin
            CommonIntegrationPlugin commonIntegration = new CommonIntegrationPlugin(serviceProvider);

            // Login and get access token
            string accessToken = commonIntegration.Login().Result;

            // Send case note details to the API
            ApiResponse response = commonIntegration.PostCaseNote(caseToken, noteDetails, accessToken).Result;
            if (!response.IsError)
            {
                response.Content = E2eeCaseDecryption.Decrypt(serviceProvider, response.Content, PayloadType.Note, parentReference);
            }
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
