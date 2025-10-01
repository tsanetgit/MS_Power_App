using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;

public class CommonCasePlugin
{
    public void ProcessCaseResponse(IOrganizationService service, ITracingService tracingService, string caseResponseJson, Guid tsanetcaseId)
    {
        try
        {

            tracingService.Trace("Starting to process case response JSON");

            // Deserialize the case response JSON
            var caseResponse = JsonConvert.DeserializeObject<Case>(caseResponseJson);

            if (caseResponse == null)
            {
                throw new InvalidPluginExecutionException("Failed to deserialize case response JSON.");
            }

            // Update the tsanetcase record with the JSON content - commented out, because of UX limitations
            //Entity tsanetcase = new Entity("ap_tsanetcase");
            //tsanetcase.Id = tsanetcaseId;
            //tsanetcase["ap_formjson"] = caseResponseJson;
            //service.Update(tsanetcase);
            //tracingService.Trace("Updated tsanetcase record with form JSON");

            // Process case notes
            ProcessCaseNotes(service, tracingService, caseResponse.CaseNotes, tsanetcaseId);
            tracingService.Trace("Processed case notes");

            // Process case responses
            ProcessCaseResponses(service, tracingService, caseResponse.CaseResponses, tsanetcaseId, caseResponse);
            tracingService.Trace("Processed case responses");

        }
        catch (Exception ex)
        {
            tracingService.Trace($"Exception: {ex.Message}");
            throw new InvalidPluginExecutionException($"An error occurred in the UpdateCasePlugin: {ex.Message}", ex);
        }
    }

    private void ProcessCaseNotes(IOrganizationService service, ITracingService tracingService, List<CaseNote> caseNotes, Guid tsanetcaseId)
    {
        if (caseNotes == null || caseNotes.Count == 0)
        {
            tracingService.Trace("No case notes to process");
            return;
        }

        foreach (var note in caseNotes)
        {
            tracingService.Trace($"Processing case note: {note.Id}");
            
            // Check if the note exists using alternate key (ap_tsanotecode)
            var query = new QueryExpression("ap_tsanetnote");
            query.ColumnSet = new ColumnSet("ap_tsanetnoteid");
            query.Criteria.AddCondition("ap_tsanotecode", ConditionOperator.Equal, note.Id.ToString());
            
            var existingNotes = service.RetrieveMultiple(query);
            
            Entity noteEntity;
            bool isNew = existingNotes.Entities.Count == 0;
            
            if (isNew)
            {
                noteEntity = new Entity("ap_tsanetnote");
                tracingService.Trace($"Creating new note with ID: {note.Id}");

                // Set note properties
                noteEntity["ap_name"] = note.Summary;
                noteEntity["ap_priority"] = GetPriorityValue(note.Priority);
                noteEntity["ap_description"] = note.Description;
                noteEntity["ap_creatoremail"] = note.CreatorEmail;
                noteEntity["ap_creatorname"] = note.CreatorName;
                noteEntity["ap_source"] = new OptionSetValue(120950001); // External source
                noteEntity["ap_tsanotecode"] = note.Id.ToString();

                if (note.CreatedAt != DateTime.MinValue)
                {
                    noteEntity["createdon"] = note.CreatedAt;
                }

                // Link to parent record
                noteEntity["ap_tsanetcaseid"] = new EntityReference("ap_tsanetcase", tsanetcaseId);
                service.Create(noteEntity);
            }
            else
            {
                noteEntity = existingNotes.Entities[0];
                tracingService.Trace($"Existing note with ID: {note.Id}");
            }
        }
    }

    private void ProcessCaseResponses(IOrganizationService service, ITracingService tracingService, List<CaseResponse> caseResponses, Guid tsanetcaseId, Case caseData)
    {
        if (caseResponses == null || caseResponses.Count == 0)
        {
            tracingService.Trace("No case responses to process");
            return;
        }
        
        // Get the ap_direction value from the parent tsanetcase
        Entity tsanetcase = service.Retrieve("ap_tsanetcase", tsanetcaseId, new ColumnSet("ap_direction"));
        int? direction = tsanetcase.Contains("ap_direction") && tsanetcase["ap_direction"] is OptionSetValue directionValue ? 
            directionValue.Value : (int?)null;
        
        foreach (var response in caseResponses)
        {
            tracingService.Trace($"Processing case response: {response.Id}");
            
            // Check if the response exists
            var query = new QueryExpression("ap_tsanetresponse");
            query.ColumnSet = new ColumnSet("ap_tsanetresponseid");
            query.Criteria.AddCondition("ap_tsaresponsecode", ConditionOperator.Equal, response.Id.ToString());
            
            var existingResponses = service.RetrieveMultiple(query);
            bool isExisting = existingResponses.Entities.Count > 0;
            
            // For existing records:
            // Only patch if ap_direction is 1 (outbound) and type is approval
            if (isExisting)
            {
                if (direction == 1 && response.Type.ToLower() == "approval")
                {
                    tracingService.Trace($"Updating existing response with ID: {response.Id}");
                    
                    Entity responseEntity = existingResponses.Entities[0];
                    
                    // Update fields
                    responseEntity["ap_type"] = GetResponseTypeValue(response.Type);
                    responseEntity["ap_tsaresponsecode"] = response.Id.ToString();
                    responseEntity["ap_engineername"] = response.EngineerName;
                    responseEntity["ap_engineerphone"] = response.EngineerPhone;
                    responseEntity["ap_engineeremail"] = response.EngineerEmail;
                    responseEntity["ap_internalcasenumber"] = response.CaseNumber;
                    responseEntity["ap_description"] = response.NextSteps;
                    
                    if (response.CreatedAt != DateTime.MinValue)
                    {
                        responseEntity["overriddencreatedon"] = response.CreatedAt;
                    }
                    
                    responseEntity["ap_source"] = new OptionSetValue(120950001); // External source
                    
                    service.Update(responseEntity);
                }
                else
                {
                    tracingService.Trace($"Skipping update for response: {response.Id} - not an outbound approval");
                }
            }
            // For new records: always create
            else
            {
                tracingService.Trace($"Creating new response with ID: {response.Id}");
                
                Entity responseEntity = new Entity("ap_tsanetresponse");
                
                // Set response properties
                responseEntity["ap_type"] = GetResponseTypeValue(response.Type);
                responseEntity["ap_tsaresponsecode"] = response.Id.ToString();
                responseEntity["ap_engineername"] = response.EngineerName;
                responseEntity["ap_engineerphone"] = response.EngineerPhone;
                responseEntity["ap_engineeremail"] = response.EngineerEmail;
                responseEntity["ap_internalcasenumber"] = response.CaseNumber;
                responseEntity["ap_description"] = response.NextSteps;
                
                if (response.CreatedAt != DateTime.MinValue)
                {
                    responseEntity["overriddencreatedon"] = response.CreatedAt;
                }
                
                responseEntity["ap_source"] = new OptionSetValue(120950001); // External source
                responseEntity["ap_tsanetcaseid"] = new EntityReference("ap_tsanetcase", tsanetcaseId);
                
                service.Create(responseEntity);
            }
        }
    }

    private OptionSetValue GetPriorityValue(string priority)
    {
        if (string.IsNullOrEmpty(priority)) return new OptionSetValue(3); // Default to LOW
        
        switch (priority.ToUpper())
        {
            case "HIGH":
                return new OptionSetValue(1);
            case "MEDIUM":
                return new OptionSetValue(2);
            case "LOW":
            default:
                return new OptionSetValue(3);
        }
    }

    private OptionSetValue GetResponseTypeValue(string type)
    {
        if (string.IsNullOrEmpty(type)) return new OptionSetValue(1); // Default to APPROVAL
        
        switch (type.ToUpper())
        {
            case "REJECTION":
                return new OptionSetValue(0);
            case "APPROVAL":
                return new OptionSetValue(1);
            case "INFORMATION_REQUEST":
                return new OptionSetValue(2);
            case "INFORMATION_RESPONSE":
                return new OptionSetValue(3);
            case "CLOSURE":
                return new OptionSetValue(4);
            case "UPDATE_APPROVAL":
                return new OptionSetValue(5);
            default:
                return new OptionSetValue(1);
        }
    }
}