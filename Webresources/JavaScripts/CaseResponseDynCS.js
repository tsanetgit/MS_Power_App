"use strict";
function onLoadDynCS(executionContext) {
    var formContext = executionContext.getFormContext();

    formContext.getAttribute("ap_incidentid").addOnChange(handleCreateCaseVisibility);
    formContext.getAttribute("ap_createcase").addOnChange(toggleCreateCaseSection);

    var apType = formContext.getAttribute("ap_type").getValue();
    
    // Only proceed for types 1 and 5
    if (apType === 1 || apType === 5) {
        // Show incident field
        formContext.getControl("ap_incidentid").setVisible(true);
        
        // Get the related case ID record
        var tsacaseId = formContext.getAttribute("ap_tsanetcaseid").getValue();
        if (tsacaseId) {
            // Get the record ID from the lookup
            var tsacaseEntityId = tsacaseId[0].id;
            
            // Retrieve the form JSON from the related record
            Xrm.WebApi.retrieveRecord("ap_tsanetcase", tsacaseEntityId, "?$select=ap_formjson,_ap_caseid_value").then(
                function success(result) {
                    if (result._ap_caseid_value) {
                        // Use the existing ap_caseid instead of searching
                        var lookupValue = [{
                            id: result._ap_caseid_value,
                            name: result["_ap_caseid_value@OData.Community.Display.V1.FormattedValue"],
                            entityType: "incident"
                        }];
                        formContext.getAttribute("ap_incidentid").setValue(lookupValue);
                        handleCreateCaseVisibility(executionContext);
                    } 
                    else if (result.ap_formjson) {
                        var formJson = JSON.parse(result.ap_formjson);
                        var customerCaseNumber = null;
                        
                        // Extract customer case number from the JSON
                        if (formJson && formJson.customFields) {
                            for (var i = 0; i < formJson.customFields.length; i++) {
                                var field = formJson.customFields[i];
                                if (field.fieldName && field.fieldName.indexOf("Customer Case #") !== -1) {
                                    customerCaseNumber = field.value;
                                    break;
                                }
                            }
                        }
                        
                        if (customerCaseNumber) {
                            formContext.getAttribute("ap_customercasenumber").setValue(customerCaseNumber);

                            // Search for incident with matching title or ticket number
                            var fetchXml = `
                                <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
                                  <entity name="incident">
                                    <attribute name="incidentid" />
                                    <attribute name="title" />
                                    <attribute name="ticketnumber" />
                                    <filter type="or">
                                      <condition attribute="title" operator="eq" value="${customerCaseNumber}" />
                                      <condition attribute="ticketnumber" operator="eq" value="${customerCaseNumber}" />
                                    </filter>
                                  </entity>
                                </fetch>`;
                                
                            var encodedFetchXml = encodeURIComponent(fetchXml);
                            
                            Xrm.WebApi.retrieveMultipleRecords("incident", `?fetchXml=${encodedFetchXml}`).then(
                                function success(result) {
                                    if (result.entities && result.entities.length > 0) {
                                        // Set the incident lookup
                                        var incident = result.entities[0];
                                        var lookupValue = [{
                                            id: incident.incidentid,
                                            name: incident.title || incident.ticketnumber,
                                            entityType: "incident"
                                        }];
                                        
                                        formContext.getAttribute("ap_incidentid").setValue(lookupValue);
                                        handleCreateCaseVisibility(executionContext);
                                    } else {
                                        // No incident found
                                        handleCreateCaseVisibility(executionContext);
                                    }
                                },
                                function error(error) {
                                    handleCreateCaseVisibility(executionContext);
                                }
                            );
                        } else {
                            // No customer case number found
                            handleCreateCaseVisibility(executionContext);
                        }
                    } else {
                        // No form JSON found
                        handleCreateCaseVisibility(executionContext);
                    }
                },
                function error(error) {
                    handleCreateCaseVisibility(executionContext);
                }
            );
        } else {
            // No case ID provided
            handleCreateCaseVisibility(executionContext);
        }
    }
}

function handleCreateCaseVisibility(executionContext) {
    var formContext = executionContext.getFormContext();
        // If incident isn't populated, show create case field
    var incidentValue = formContext.getAttribute("ap_incidentid").getValue();
    if (!incidentValue || !incidentValue.length) {
        // Show create case field
        formContext.getControl("ap_createcase").setVisible(true);
    } else {
        // Hide create case field if incident is populated
        formContext.getControl("ap_createcase").setVisible(false);
    }
}

function toggleCreateCaseSection(executionContext) {
    var formContext = executionContext.getFormContext();
    var createCase = formContext.getAttribute("ap_createcase").getValue();
    var createCaseSection = formContext.ui.tabs.get("tab_1").sections.get("createcasesection");
    
    if (createCase === true) {
        // Show section and make customer field required
        createCaseSection.setVisible(true);
        formContext.getAttribute("ap_customerid").setRequiredLevel("required");
    } else {
        // Hide section and make customer field not required
        createCaseSection.setVisible(false);
        formContext.getAttribute("ap_customerid").setRequiredLevel("none");
    }
}
