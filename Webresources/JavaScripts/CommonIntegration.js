// Get Case
function getCase(formContext) {
    const internalCaseNumber = formContext.getAttribute("ap_submittercasenumber").getValue();

    if (!internalCaseNumber) {
        showError(formContext, "Internal case number is required.");
        return;
    }
    Xrm.Utility.showProgressIndicator("Retrieving case details...");

    const parameters = {
        InternalCaseNumber: internalCaseNumber
    };

    const request = {
        InternalCaseNumber: parameters.InternalCaseNumber,
        getMetadata: function () {
            return {
                boundParameter: null,
                parameterTypes: {
                    "InternalCaseNumber": { typeName: "Edm.String", structuralProperty: 1 }
                },
                operationType: 0,
                operationName: "ap_GetCase"
            };
        }
    };

    Xrm.WebApi.online.execute(request).then(
        function success(result) {
            if (result.ok) {
                result.json().then(function (response) {
                    if (!response.IsError) {
                        var formJson = response.GetCaseResponse;
                        var formResponse = JSON.parse(formJson);
                        console.log(formResponse);
                        saveToFormField("ap_formjson", formResponse, formContext);  // Save JSON
                        // Process caseNotes array to patch each note
                        let caseNotes = formResponse.caseNotes;
                        let updateNotePromises = caseNotes.map(note => {
                            // Step 1: Prepare the update data for each note
                            let updateData = {
                                ap_name: note.summary,
                                ap_priority: getPriorityValue(note.priority),
                                ap_description: note.description,
                                ap_creatoremail: note.creatorEmail,
                                ap_creatorname: note.creatorName,
                                "ap_tsanetcaseid@odata.bind": `/ap_tsanetcases(${formContext.data.entity.getId().replace("{", "").replace("}", "")})`
                            };

                            // Step 2: Construct the URL with alternate key
                            let apiUrl = Xrm.Utility.getGlobalContext().getClientUrl() +
                                `/api/data/v9.2/ap_tsanetnotes(ap_tsanotecode='${note.id}')`;

                            // Step 3: Create the PATCH request using Web API
                            return sendPatchRequest(apiUrl, updateData);
                        });

                        // Process caseResponses array to create or patch each response
                        let caseResponses = formResponse.caseResponses;
                        let updateResponsePromises = caseResponses.map(response => {
                            // Step 1: Prepare the update data for each response
                            let responseData = {
                                ap_type: getResponseTypeValue(response.type),
                                ap_tsaresponsecode: response.id.toString(),
                                ap_engineername: response.engineerName,
                                ap_engineerphone: response.engineerPhone,
                                ap_engineeremail: response.engineerEmail,
                                ap_internalcasenumber: response.caseNumber,
                                ap_description: response.nextSteps,
                                overriddencreatedon: response.createdAt,
                                ap_source: 120950001,
                                "ap_tsanetcaseid@odata.bind": `/ap_tsanetcases(${formContext.data.entity.getId().replace("{", "").replace("}", "")})`
                            };

                            // Step 2: Check if the record exists and create or patch accordingly
                            return checkRecordExists(response.id).then(recordId => {
                                if (recordId) {
                                    // Only patch if ap_direction is 1 and type is approval
                                    if (formContext.getAttribute("ap_direction").getValue() === 1 && response.type.toLowerCase() === 'approval') {
                                        return Xrm.WebApi.updateRecord("ap_tsanetresponse", recordId, responseData);
                                    } else {
                                        return Promise.resolve(); // Skip patching
                                    }
                                } else {
                                    return Xrm.WebApi.createRecord("ap_tsanetresponse", responseData);
                                }
                            });
                        });

                        // Execute all PATCH/CREATE requests and handle final actions
                        Promise.all([...updateNotePromises, ...updateResponsePromises])
                            .then(() => {
                                Xrm.Utility.closeProgressIndicator();
                                formContext.ui.setFormNotification("Successfully updated!", "INFO", "success");
                                // Refresh the notes subgrid
                                //formContext.getControl("notesubgrid").refresh();
                                refreshReadOnlyForm(formContext);
                            })
                            .catch(error => {
                                Xrm.Utility.closeProgressIndicator();
                                console.error(error.message);
                                showError(formContext, "Error updating case notes or responses: " + error.message);
                            });

                    } else {
                        Xrm.Utility.closeProgressIndicator();
                        var error = JSON.parse(response.GetCaseResponse);
                        showError(formContext, error.message);
                    }
                });
            }
        },
        function (error) {
            Xrm.Utility.closeProgressIndicator();
            console.error(error.message);
            showError(formContext, error.message);
        }
    );
}

// Helper function to send PATCH request
function sendPatchRequest(url, data) {
    return new Promise((resolve, reject) => {
        const req = new XMLHttpRequest();
        req.open("PATCH", url, true);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");

        req.onreadystatechange = function () {
            if (req.readyState === 4) {
                if (req.status === 204) {
                    resolve(); // Success
                } else {
                    reject(new Error(`Failed to update record. Status: ${req.status}, Error: ${req.responseText}`));
                }
            }
        };

        req.send(JSON.stringify(data));
    });
}

// Helper function to check if a record exists
function checkRecordExists(responseCode) {
    const query = `?$filter=ap_tsaresponsecode eq '${responseCode}'&$select=ap_tsanetresponseid`;
    return Xrm.WebApi.retrieveMultipleRecords("ap_tsanetresponse", query).then(
        function success(result) {
            if (result.entities.length > 0) {
                return result.entities[0].ap_tsanetresponseid; // Return the record ID
            } else {
                return null; // Record does not exist
            }
        },
        function (error) {
            throw new Error(`Failed to check record existence. Status: ${error.status}, Error: ${error.message}`);
        }
    );
}

// Get company
function getCompanyDetails(companyName) {
    return new Promise(function (resolve, reject) {
        //Xrm.Utility.showProgressIndicator("Retrieving company details...");
        const parameters = { CompanyName: companyName };
        const request = {
            CompanyName: parameters.CompanyName,
            getMetadata: function () {
                return {
                    boundParameter: null,
                    parameterTypes: {
                        "CompanyName": { typeName: "Edm.String", structuralProperty: 1 }
                    },
                    operationType: 0,
                    operationName: "ap_GetCompanyByName"
                };
            }

        };

        Xrm.WebApi.online.execute(request).then(
            function success(result) {
                if (result.ok) {
                    result.json().then(
                        function (response) {
                            const companyDetailsJson = response.CompanyDetails;
                            const companyDetails = JSON.parse(companyDetailsJson);
                            //Xrm.Utility.closeProgressIndicator();
                            resolve(companyDetails);
                        });
                } else {
                    //Xrm.Utility.closeProgressIndicator();
                    reject(new Error("No result from company search"));
                }
            },
            function (error) {
                //Xrm.Utility.closeProgressIndicator();
                reject(error);
            }
        );
    });
}

// getFormByCompany function
function getFormByCompany(companyId, formContext) {

    Xrm.Utility.showProgressIndicator("Retrieving form details...");
    const parameters = {
        CompanyId: companyId
    };

    // Custom action call
    const request = {
        CompanyId: parameters.CompanyId, 
        getMetadata: function () {
            return {
                boundParameter: null, 
                parameterTypes: {
                    "CompanyId": { typeName: "Edm.Int32", structuralProperty: 1 } 
                },
                operationType: 0, 
                operationName: "ap_GetFormByCompany"
            };
        }
    };

    Xrm.WebApi.online.execute(request).then(
        function success(result) {
            if (result.ok) {
                result.json().then(function (response) {
                    if (!response.IsError) {
                        var formJson = response.FormDetails;
                        var formDetails = JSON.parse(formJson);
                        console.log(formDetails);
                        Xrm.Utility.closeProgressIndicator();
                        displayDynamicForm(formDetails, formContext);
                    }
                    else {
                        Xrm.Utility.closeProgressIndicator();
                        var error = JSON.parse(response.ErrorMessage);
                        showError(formContext, error.message);
                    }
                });
            }
        },
        function (error) {
            Xrm.Utility.closeProgressIndicator();
            console.error(error.message); // Log any errors in the console
            showError(formContext, error.message);
        }
    );
}

function postCase(submissionData, formContext) {
    disableButton(true, "WebResource_casecreate");
    // Convert the submissionData object to a JSON string
    const submissionDataString = JSON.stringify(submissionData);
    Xrm.Utility.showProgressIndicator("Sending case...");
    const parameters = {
        CaseDetails: submissionDataString
    };

    // Custom action call
    const request = {
        CaseDetails: parameters.CaseDetails, 
        getMetadata: function () {
            return {
                boundParameter: null, // No entity bound
                parameterTypes: {
                    "CaseDetails": { typeName: "Edm.String", structuralProperty: 1 } 
                },
                operationType: 0, 
                operationName: "ap_PostCase"
            };
        }
    };

    Xrm.WebApi.online.execute(request).then(
        function success(result) {
            if (result.ok) {
                result.json().then(function (response) {
                    if (!response.IsError) {
                        var formJson = response.PostCaseResponse;
                        var formResponse = JSON.parse(formJson);
                        console.log(formResponse);
                        //save data
                        formContext.getAttribute("ap_name").setValue(formResponse.id.toString());
                        formContext.getAttribute("ap_submittercasenumber").setValue(formResponse.submitterCaseNumber.toString());
                        saveToFormField("ap_formjson", formResponse, formContext);  // Save JSON
                        Xrm.Utility.closeProgressIndicator();
                        formContext.ui.setFormNotification("Successfully created!", "INFO", "success");
                        // Save the record
                        formContext.data.entity.save();
                    }
                    else {
                        Xrm.Utility.closeProgressIndicator();
                        var error = JSON.parse(response.PostCaseResponse);
                        showError(formContext, error.message);
                        disableButton(false, "WebResource_casecreate");
                    }
                });
            }
        },
        function (error) {
            Xrm.Utility.closeProgressIndicator();
            console.error(error.message);
            showError(formContext, error.message);
            disableButton(false, "WebResource_casecreate");
        }
    );
}

// Get Cases
function getCases() {
    Xrm.Utility.showProgressIndicator("Case refresh started");

    return new Promise(function (resolve, reject) {
        const parameters = {};
        const request = {
            getMetadata: function () {
                return {
                    boundParameter: null,
                    parameterTypes: {
                    },
                    operationType: 0,
                    operationName: "ap_GetCases"
                };
            }
        };

        Xrm.WebApi.online.execute(request).then(
            function success(result) {
                if (result.ok) {
                    Xrm.Utility.closeProgressIndicator();
                    Xrm.Utility.alertDialog("Cases refresh started successfully. It usually takes 30-60 seconds to refresh data entirely.");
                } else {
                    Xrm.Utility.closeProgressIndicator();
                    Xrm.Utility.alertDialog("Error - can't refresh cases");
                    reject(new Error("Error - can't refresh cases"));
                }
            },
            function (error) {
                Xrm.Utility.closeProgressIndicator();
                Xrm.Utility.alertDialog("Error - can't refresh cases: " + error.message);
                reject(error);
            }
        );
    });
}
