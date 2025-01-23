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
            Entity caseEntity = service.Retrieve(caseReference.LogicalName, caseReference.Id, new ColumnSet("statecode"));

            if (!caseEntity.Contains("statecode"))
            {
                throw new InvalidPluginExecutionException("statecode is missing in the related tsanetcase record.");
            }

            // Check if the statecode is 0 (active)
            OptionSetValue stateCode = caseEntity.GetAttributeValue<OptionSetValue>("statecode");
            OptionSetValue source = entity.GetAttributeValue<OptionSetValue>("ap_source");
            if (stateCode.Value == 1 && source.Value == 120950000)
            {
                throw new InvalidPluginExecutionException("You can't create notes for inactive cases.");
            }
        }
    }
}
