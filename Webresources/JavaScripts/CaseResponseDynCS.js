"use strict";
function onLoadDynCS(executionContext) {
    var formContext = executionContext.getFormContext();

    formContext.getControl("ap_internalcasenumber").setDisabled(true);

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

            // First retrieve mapping configurations
            retrieveCaseMappingConfigurations().then(function (mappingConfigs) {
                // Now retrieve the form JSON from the related record
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

                            // Process all mappings from configuration
                            processMappingsFromJson(formContext, formJson, mappingConfigs).then(function () {
                                handleCreateCaseVisibility(executionContext);
                            }).catch(function (error) {
                                handleCreateCaseVisibility(executionContext);
                            });
                        } else {
                            // No form JSON found
                            handleCreateCaseVisibility(executionContext);
                        }
                    },
                    function error(error) {
                        handleCreateCaseVisibility(executionContext);
                    }
                );
            }).catch(function (error) {
                handleCreateCaseVisibility(executionContext);
            });
        } else {
            // No case ID provided
            handleCreateCaseVisibility(executionContext);
        }
    }
}

function retrieveCaseMappingConfigurations() {
    return new Promise(function (resolve, reject) {
        var fetchXml = `
            <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
              <entity name="ap_casemapping">
                <attribute name="ap_casemappingid" />
                <attribute name="ap_name" />
                <attribute name="ap_sourcejsonpath" />
                <attribute name="ap_targetattribute" />
                <attribute name="ap_attributetype" />
                <attribute name="ap_lookupentityname" />
                <attribute name="ap_lookupsearchattribute" />
                <order attribute="ap_name" descending="false" />
                <filter>
                   <condition attribute="ap_mappingtype" operator="eq" value="120950000" />
                </filter>
              </entity>
            </fetch>`;

        var encodedFetchXml = encodeURIComponent(fetchXml);

        Xrm.WebApi.retrieveMultipleRecords("ap_casemapping", `?fetchXml=${encodedFetchXml}`).then(
            function success(result) {
                if (result.entities && result.entities.length > 0) {
                    resolve(result.entities);
                } else {
                    resolve([]);
                }
            },
            function error(error) {
                reject(error);
            }
        );
    });
}

function processMappingsFromJson(formContext, formJson, mappingConfigs) {
    return new Promise(function (resolve, reject) {
        // Create an array to track all promises
        var mappingPromises = [];

        mappingConfigs.forEach(function (mapping) {
            var sourceValue = extractValueFromJson(formJson, mapping.ap_sourcejsonpath);

            if (sourceValue !== null) {
                var mappingPromise = processMapping(formContext, mapping, sourceValue);
                mappingPromises.push(mappingPromise);
            }
        });

        // Wait for all mappings to complete
        Promise.all(mappingPromises).then(function () {
            resolve();
        }).catch(function (error) {
            reject(error);
        });
    });
}

function extractValueFromJson(formJson, sourcePath) {
    // Handle paths like "customFields.Customer Case #"
    var pathParts = sourcePath.split('.');

    if (pathParts.length > 0) {
        var currentObj = formJson;

        // Navigate to the specified object section (like customFields)
        if (pathParts[0] && currentObj[pathParts[0]]) {
            currentObj = currentObj[pathParts[0]];

            // If we're looking for a specific named field (like "Customer Case #")
            if (pathParts.length > 1 && Array.isArray(currentObj)) {
                var fieldName = pathParts[1];
                for (var i = 0; i < currentObj.length; i++) {
                    var field = currentObj[i];
                    if (field.fieldName && field.fieldName.indexOf(fieldName) !== -1) {
                        return field.value;
                    }
                }
            }
        }
    }

    return null;
}

function processMapping(formContext, mapping, sourceValue) {
    return new Promise(function (resolve, reject) {
        try {
            // Check if sourceValue is empty (null, undefined, empty string, or just whitespace)
            if (sourceValue === null || sourceValue === undefined ||
                (typeof sourceValue === "string" && sourceValue.trim() === "")) {
                // Skip processing if empty
                resolve();
                return;
            }

            var attributeType = mapping.ap_attributetype;
            var targetAttribute = mapping.ap_targetattribute;

            switch (attributeType) {
                case 3:
                case 4:
                case 1:
                    // Direct assignment for simple types
                    formContext.getAttribute(targetAttribute).setValue(sourceValue);
                    resolve();
                    break;

                case 2:
                    // Handle lookup fields by searching for the entity
                    var entityName = mapping.ap_lookupentityname;
                    var searchAttribute = mapping.ap_lookupsearchattribute || "name";

                    // Create a FetchXML to search for the entity
                    var fetchXml = `
                        <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
                          <entity name="${entityName}">
                            <attribute name="${entityName}id" />
                            <attribute name="${searchAttribute}" />
                            <filter>
                              <condition attribute="${searchAttribute}" operator="eq" value="${sourceValue}" />
                            </filter>
                          </entity>
                        </fetch>`;

                    var encodedFetchXml = encodeURIComponent(fetchXml);

                    Xrm.WebApi.retrieveMultipleRecords(entityName, `?fetchXml=${encodedFetchXml}`).then(
                        function success(result) {
                            if (result.entities && result.entities.length > 0) {
                                var entity = result.entities[0];
                                var lookupValue = [{
                                    id: entity[`${entityName}id`],
                                    name: entity[searchAttribute],
                                    entityType: entityName
                                }];

                                formContext.getAttribute(targetAttribute).setValue(lookupValue);
                            }
                            resolve();
                        },
                        function error(error) {
                            resolve(); // Resolve anyway to continue with other mappings
                        }
                    );
                    break;

                default:
                    resolve();
                    break;
            }
        } catch (e) {
            resolve(); // Resolve anyway to continue with other mappings
        }
    });
}

function handleCreateCaseVisibility(executionContext) {
    var formContext = executionContext.getFormContext();
        // If incident isn't populated, show create case field
    var incidentValue = formContext.getAttribute("ap_incidentid").getValue();
    if (!incidentValue || !incidentValue.length) {
        // Show create case field
        formContext.getControl("ap_createcase").setVisible(true);
        formContext.getAttribute("ap_internalcasenumber").setValue(null);
    } else {
        // Hide create case field if incident is populated
        formContext.getControl("ap_createcase").setVisible(false);
        setInternalCaseNumberFromIncident(formContext, incidentValue[0].id);
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

function setInternalCaseNumberFromIncident(formContext, incidentId) {
    // Clean the GUID (remove curly braces if present)
    var cleanIncidentId = incidentId.replace(/[{}]/g, '');

    // Retrieve the incident to get the ticketnumber
    Xrm.WebApi.retrieveRecord("incident", cleanIncidentId, "?$select=ticketnumber").then(
        function success(result) {
            if (result.ticketnumber) {
                formContext.getAttribute("ap_internalcasenumber").setValue(result.ticketnumber);
            }
        },
        function error(error) {
            console.log("Error retrieving incident ticketnumber: " + error.message);
        }
    );
}