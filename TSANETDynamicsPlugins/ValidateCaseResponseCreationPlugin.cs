using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace TSANetDynamicsPlugins
{
    public class ValidateCaseResponseCreationPlugin : IPlugin
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

                // Check if the entity contains the parent case reference and source
                if (!entity.Contains("ap_tsanetcaseid") || !entity.Contains("ap_source") | !entity.Contains("ap_type"))
                {
                    throw new InvalidPluginExecutionException("ap_tsanetcaseid or ap_source or ap_type is missing.");
                }

                // Get TSA Net case first
                EntityReference caseReference = (EntityReference)entity["ap_tsanetcaseid"];
                Entity caseEntity = service.Retrieve(caseReference.LogicalName, caseReference.Id, new ColumnSet("ap_caseid"));
                OptionSetValue source = entity.GetAttributeValue<OptionSetValue>("ap_source");
                OptionSetValue type = entity.GetAttributeValue<OptionSetValue>("ap_type");

                // Check if source is dynamics
                if (source.Value == 120950000)
                {
                    // Validate if the case ID is set
                    if (!caseEntity.Contains("ap_caseid") && type.Value == 1)
                    {
                        throw new InvalidPluginExecutionException("You must assign the case first");
                    }
                }
            }
        }
    }
}