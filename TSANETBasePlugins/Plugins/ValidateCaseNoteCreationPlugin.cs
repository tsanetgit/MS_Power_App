using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

public class ValidateCaseNoteCreationPlugin : IPlugin
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

            // Check if the entity contains the parent case reference
            if (!entity.Contains("ap_tsanetcaseid") || !entity.Contains("ap_source"))
            {
                throw new InvalidPluginExecutionException("ap_tsanetcaseid or ap_source is missing.");
            }

            EntityReference caseReference = (EntityReference)entity["ap_tsanetcaseid"];
            Entity caseEntity = service.Retrieve(caseReference.LogicalName, caseReference.Id, new ColumnSet("statecode", "statuscode"));
            OptionSetValue stateCode = caseEntity.GetAttributeValue<OptionSetValue>("statecode");
            OptionSetValue statusCode = caseEntity.GetAttributeValue<OptionSetValue>("statuscode");
            OptionSetValue source = entity.GetAttributeValue<OptionSetValue>("ap_source");

            // Check if the it is allowed to create notes for the case status
            if (source.Value == 120950000)
            {
                if (stateCode.Value == 1)
                {
                    throw new InvalidPluginExecutionException("You can't create notes for inactive cases.");
                }
                else if (statusCode.Value == 1)
                {
                    throw new InvalidPluginExecutionException("You can't create notes for cases with status 'Open'.");
                }
            }
        }
    }
}
