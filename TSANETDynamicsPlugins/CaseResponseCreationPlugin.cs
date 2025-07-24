using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.IdentityModel.Metadata;

namespace TSANetDynamicsPlugins
{
    public class CaseResponseCreatePlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("CaseResponseCreatePlugin: Execution started");

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity caseResponse = (Entity)context.InputParameters["Target"];
                tracingService.Trace("CaseResponseCreatePlugin: Retrieved case response entity");

                OptionSetValue source = caseResponse.GetAttributeValue<OptionSetValue>("ap_source");

                if (source.Value != 120950000)
                {
                    tracingService.Trace("CaseResponseCreatePlugin: Case response is not created from Dynamics");
                    return; // Exit if source is not Dynamics
                }


                if (caseResponse.Contains("ap_casecreate") && caseResponse.GetAttributeValue<bool>("ap_casecreate") == true &&
                    (!caseResponse.Contains("ap_incidentid") || caseResponse["ap_incidentid"] == null))
                {
                    tracingService.Trace("CaseResponseCreatePlugin: Case creation required");                    

                    // Create new incident with prefilled data
                    if (caseResponse.Contains("ap_customerid") && caseResponse.Contains("ap_customercasenumber"))
                    {
                        tracingService.Trace("CaseResponseCreatePlugin: Creating new incident");
                        Entity incident = new Entity("incident");

                        // Set customerid and title
                        incident["customerid"] = caseResponse["ap_customerid"];
                        incident["title"] = caseResponse["ap_customercasenumber"];

                        // Create the incident
                        Guid incidentId = service.Create(incident);
                        tracingService.Trace($"CaseResponseCreatePlugin: Created incident with ID: {incidentId}");

                        // Step 3: Update field ap_incidentid with created incident
                        Entity updateCaseResponse = new Entity(caseResponse.LogicalName);
                        updateCaseResponse.Id = caseResponse.Id;
                        updateCaseResponse["ap_incidentid"] = new EntityReference("incident", incidentId);
                        service.Update(updateCaseResponse);
                        tracingService.Trace("CaseResponseCreatePlugin: Updated case response with incident reference");

                        // Update the caseResponse entity with the incident ID for use in the next step
                        caseResponse["ap_incidentid"] = new EntityReference("incident", incidentId);
                    }
                    else
                    {
                        tracingService.Trace("CaseResponseCreatePlugin: Missing required fields for incident creation");
                        throw new InvalidPluginExecutionException("Customer or Customer # is missing");
                    }
                }

                // Step 4: Update parent case if ap_incidentid is not empty
                if (caseResponse.Contains("ap_incidentid") && caseResponse["ap_incidentid"] != null &&
                    caseResponse.Contains("ap_tsanetcaseid") && caseResponse["ap_tsanetcaseid"] != null)
                {
                    tracingService.Trace("CaseResponseCreatePlugin: Updating parent case");
                    EntityReference incidentRef = (EntityReference)caseResponse["ap_incidentid"];
                    EntityReference parentCaseRef = (EntityReference)caseResponse["ap_tsanetcaseid"];

                    Entity updateParentCase = new Entity(parentCaseRef.LogicalName);
                    updateParentCase.Id = parentCaseRef.Id;
                    updateParentCase["ap_caseid"] = incidentRef;
                    service.Update(updateParentCase);
                    tracingService.Trace("CaseResponseCreatePlugin: Updated parent case with incident reference");
                }
            }

            tracingService.Trace("CaseResponseCreatePlugin: Execution completed");
        }
    }
}