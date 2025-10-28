using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IdentityModel.Metadata;
using System.Linq;

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
                tracingService.Trace($"CaseResponseCreatePlugin: Source is: " + source.Value);

                if (source.Value != 120950000)
                {
                    tracingService.Trace("CaseResponseCreatePlugin: Case response is not created from Dynamics");
                    return; // Exit if source is not Dynamics
                }
                tracingService.Trace($"CaseResponseCreatePlugin: Source is: Dynamics");

                if (caseResponse.GetAttributeValue<bool>("ap_createcase"))
                {
                    tracingService.Trace($"CaseResponseCreatePlugin: Starting the create case");

                    // Retrieve case mapping configurations
                    var caseMappings = RetrieveCaseMappingConfigurations(service, tracingService);
                    tracingService.Trace($"CaseResponseCreatePlugin: Retrieved {caseMappings.Count} case mapping configurations");

                    // Create new incident with prefilled data from configuration
                    Entity incident = new Entity("incident");

                    // Apply mappings using the separate method
                    var mappingResult = ApplyMappings(caseResponse, incident, caseMappings, tracingService);

                    if (!mappingResult.Success)
                    {
                        throw new InvalidPluginExecutionException(mappingResult.ErrorMessage);
                    }

                    // Create the incident if we have the minimum required data
                    if (incident.Attributes.Count > 0)
                    {
                        tracingService.Trace("CaseResponseCreatePlugin: Creating new incident");
                        Guid incidentId = service.Create(incident);
                        tracingService.Trace($"CaseResponseCreatePlugin: Created incident with ID: {incidentId}");

                        // Retrieve the created incident to get the ticket number
                        Entity createdIncident = service.Retrieve("incident", incidentId, new ColumnSet("ticketnumber"));
                        string ticketNumber = createdIncident.GetAttributeValue<string>("ticketnumber");
                        tracingService.Trace($"CaseResponseCreatePlugin: Retrieved ticket number: {ticketNumber}");

                        // Set ap_incidentid and ap_internalcasenumber on the Target entity (PreOperation)
                        caseResponse["ap_incidentid"] = new EntityReference("incident", incidentId);
                        caseResponse["ap_internalcasenumber"] = ticketNumber;
                        tracingService.Trace("CaseResponseCreatePlugin: Set incident reference and ticket number on case response");

                    }
                    else
                    {
                        tracingService.Trace("CaseResponseCreatePlugin: No fields mapped for incident creation");
                        throw new InvalidPluginExecutionException("No fields configured for incident creation");
                    }
                }

                // Update parent tsacase if ap_incidentid is not empty
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

        /// <summary>
        /// Applies field mappings from source entity to target entity based on mapping configurations
        /// </summary>
        /// <param name="sourceEntity">Source entity containing the original values</param>
        /// <param name="targetEntity">Target entity where values will be mapped</param>
        /// <param name="mappingConfigs">List of mapping configuration entities</param>
        /// <param name="tracingService">Tracing service for logging</param>
        /// <returns>MappingResult containing success status and error message if any</returns>
        private MappingResult ApplyMappings(Entity sourceEntity, Entity targetEntity, List<Entity> mappingConfigs, ITracingService tracingService)
        {
            bool hasRequiredFields = true;
            string missingFields = "";

            foreach (var mapping in mappingConfigs)
            {
                string sourceAttribute = mapping.GetAttributeValue<string>("ap_sourcejsonpath");
                string targetAttribute = mapping.GetAttributeValue<string>("ap_targetattribute");
                bool isRequired = mapping.GetAttributeValue<bool>("ap_isrequired");
                int attributeType = mapping.GetAttributeValue<OptionSetValue>("ap_attributetype")?.Value ?? 1; // Default to direct mapping

                tracingService.Trace($"Mapping: {sourceAttribute} -> {targetAttribute} (Type: {attributeType})");

                if (string.IsNullOrEmpty(sourceAttribute) || string.IsNullOrEmpty(targetAttribute))
                {
                    tracingService.Trace("Invalid mapping configuration: source or target attribute is empty");
                    continue;
                }

                // Check if source attribute exists in source entity
                if (sourceEntity.Contains(sourceAttribute))
                {
                    object sourceValue = sourceEntity[sourceAttribute];

                    // Handle different attribute types
                    switch (attributeType)
                    {
                        case 1: // Direct mapping (string, number, boolean, etc.)
                            targetEntity[targetAttribute] = sourceValue;
                            tracingService.Trace($"Direct mapped {sourceAttribute} to {targetAttribute}");
                            break;

                        case 2: // Lookup field mapping (EntityReference to EntityReference)
                            if (sourceValue is EntityReference sourceRef)
                            {
                                targetEntity[targetAttribute] = sourceRef;
                                tracingService.Trace($"Mapped lookup {sourceAttribute} to {targetAttribute}");
                            }
                            break;

                        case 3: // Option set mapping
                            if (sourceValue is OptionSetValue optionSetValue)
                            {
                                targetEntity[targetAttribute] = new OptionSetValue(optionSetValue.Value);
                                tracingService.Trace($"Mapped option set {sourceAttribute} to {targetAttribute}");
                            }
                            break;

                        case 4: // Money field mapping
                            if (sourceValue is Money money)
                            {
                                targetEntity[targetAttribute] = money;
                                tracingService.Trace($"Mapped money {sourceAttribute} to {targetAttribute}");
                            }
                            break;

                        default:
                            // For any other type, try direct mapping
                            targetEntity[targetAttribute] = sourceValue;
                            tracingService.Trace($"Default mapped {sourceAttribute} to {targetAttribute}");
                            break;
                    }
                }
                else if (isRequired)
                {
                    hasRequiredFields = false;
                    missingFields += $"{sourceAttribute}, ";
                    tracingService.Trace($"Missing required field {sourceAttribute}");
                }
            }

            if (!hasRequiredFields)
            {
                missingFields = missingFields.TrimEnd(',', ' ');
                return new MappingResult
                {
                    Success = false,
                    ErrorMessage = $"Missing required fields: {missingFields}"
                };
            }

            return new MappingResult { Success = true };
        }

        private List<Entity> RetrieveCaseMappingConfigurations(IOrganizationService service, ITracingService tracingService)
        {
            tracingService.Trace("Retrieving case mapping configurations");

            var query = new QueryExpression("ap_casemapping")
            {
                ColumnSet = new ColumnSet(
                    "ap_casemappingid",
                    "ap_name",
                    "ap_sourcejsonpath",
                    "ap_targetattribute",
                    "ap_attributetype",
                    "ap_mappingtype",
                    "ap_isrequired"
                ),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("statecode", ConditionOperator.Equal, 0) // Active records
                    }
                },
                Orders =
                {
                    new OrderExpression("ap_name", OrderType.Ascending)
                }
            };

            var filterMappingType = new FilterExpression(LogicalOperator.Or);
            filterMappingType.AddCondition("ap_mappingtype", ConditionOperator.Equal, 120950001); // Incident Creation mapping
            query.Criteria.AddFilter(filterMappingType);

            EntityCollection results = service.RetrieveMultiple(query);
            tracingService.Trace($"Retrieved {results.Entities.Count} mapping records");
            return results.Entities.ToList();
        }

        /// <summary>
        /// Class to hold the result of the mapping operation
        /// </summary>
        private class MappingResult
        {
            public bool Success { get; set; }
            public string ErrorMessage { get; set; } = string.Empty;
        }
    }
}