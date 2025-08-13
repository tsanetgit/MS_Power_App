// Constant for standard alert dialog options
const STANDARD_ALERT_OPTIONS = { height: 120, width: 260 };

// Get Case
function getCase(formContext) {
    "use strict";
    const caseToken = formContext.getAttribute("ap_tsacasetoken").getValue();

    if (!caseToken) {
        showError(formContext, "Case Token is required.");
        return;
    }
    Xrm.Utility.showProgressIndicator("Retrieving case details...");

    const parameters = {
        CaseToken: caseToken
    };

    const request = {
        CaseToken: parameters.CaseToken,
        getMetadata: function () {
            return {
                boundParameter: null,
                parameterTypes: {
                    "CaseToken": { typeName: "Edm.String", structuralProperty: 1 }
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
                        saveToFormField("ap_formjson", formResponse, formContext);  // Save JSON
                        updateTsaCaseStatus(formContext, formResponse.status);
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
                                ap_source: 120950001,
                                createdon: note.createdAt,
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
                                showSuccess(formContext, "Successfully updated!");
                                // Refresh the notes subgrid
                                //formContext.getControl("notesubgrid").refresh();
                                refreshReadOnlyForm(formContext);
                            })
                            .catch(error => {
                                Xrm.Utility.closeProgressIndicator();
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
            showError(formContext, error.message);
        }
    );
}

// Helper function to send PATCH request
function sendPatchRequest(url, data) {
    "use strict";
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
    "use strict";
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
    "use strict";
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
    "use strict";
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
            showError(formContext, error.message);
        }
    );
}

// getFormByDepartment function
function getFormByDepartment(departmentId, formContext) {
    "use strict";
    Xrm.Utility.showProgressIndicator("Retrieving form details...");
    const parameters = {
        DepartmentId: departmentId
    };

    // Custom action call
    const request = {
        DepartmentId: parameters.DepartmentId,
        getMetadata: function () {
            return {
                boundParameter: null,
                parameterTypes: {
                    "DepartmentId": { typeName: "Edm.Int32", structuralProperty: 1 }
                },
                operationType: 0,
                operationName: "ap_GetFormByDepartment"
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
            showError(formContext, error.message);
        }
    );
}

function postCase(submissionData, formContext) {
    "use strict";
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
                        //save data
                        formContext.getAttribute("ap_name").setValue(formResponse.id.toString());
                        formContext.getAttribute("ap_submittercasenumber").setValue(formResponse.submitterCaseNumber.toString());
                        formContext.getAttribute("ap_tsacasetoken").setValue(formResponse.token.toString());
                        saveToFormField("ap_formjson", formResponse, formContext);  // Save JSON
                        Xrm.Utility.closeProgressIndicator();
                        showSuccess(formContext, "Successfully created!");
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
            showError(formContext, error.message);
            disableButton(false, "WebResource_casecreate");
        }
    );
}

// Get Cases
function getCases() {
    "use strict";
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
                    var alertStrings = { 
                        confirmButtonLabel: "OK", 
                        text: "Cases refresh started successfully. It usually takes 30-60 seconds to refresh data entirely.", 
                        title: "Success" 
                    };
                    Xrm.Navigation.openAlertDialog(alertStrings, STANDARD_ALERT_OPTIONS).then(
                        function (success) {
                            // Dialog closed successfully
                        },
                        function (error) {
                            // Error in dialog (empty handler but keeping for future use)
                        }
                    );
                } else {
                    Xrm.Utility.closeProgressIndicator();
                    var alertStrings = { 
                        confirmButtonLabel: "OK", 
                        text: "Error - can't refresh cases", 
                        title: "Error" 
                    };
                    Xrm.Navigation.openAlertDialog(alertStrings, STANDARD_ALERT_OPTIONS).then(
                        function (success) {
                            // Dialog closed successfully
                        },
                        function (error) {
                            // Error in dialog (empty handler but keeping for future use)
                        }
                    );
                    reject(new Error("Error - can't refresh cases"));
                }
            },
            function (error) {
                Xrm.Utility.closeProgressIndicator();
                var alertStrings = { 
                    confirmButtonLabel: "OK", 
                    text: "Error - can't refresh cases: " + error.message, 
                    title: "Error" 
                };
                Xrm.Navigation.openAlertDialog(alertStrings, STANDARD_ALERT_OPTIONS).then(
                    function (success) {
                        // Dialog closed successfully
                    },
                    function (error) {
                        // Error in dialog (empty handler but keeping for future use)
                    }
                );
                reject(error);
            }
        );
    });
}

// Function to create a note and upload a file
function createNoteWithFile(formContext, file) {
    "use strict";
    const caseId = formContext.data.entity.getId();

    if (!caseId) {
        showError(formContext, "No case record is available to attach the file.");
        return Promise.reject(new Error("No case record is available to attach the file."));
    }

    Xrm.Utility.showProgressIndicator("Uploading");

    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function (event) {
            const fileContent = event.target.result.split(",")[1]; // Base64 content
            const noteData = {
                "subject": "%TOPROCESS%",
                "filename": file.name,
                "documentbody": fileContent,
                "mimetype": file.type,
                "notetext": "File attached via upload.",
                "objectid_ap_tsanetcase@odata.bind": `/ap_tsanetcases(${caseId.replace(/[{}]/g, "")})`
            };

            // Create the note record
            Xrm.WebApi.createRecord("annotation", noteData).then(
                function success(result) {
                    Xrm.Utility.closeProgressIndicator();
                    showSuccess(formContext, "Success - the file is now being uploaded");

                    // Show the "Upload in progress" notification
                    showUploadInProgressNotification(formContext);

                    // Start the periodic check for annotation status
                    startAnnotationStatusCheck(formContext);

                    resolve(result);
                },
                function (error) {
                    Xrm.Utility.closeProgressIndicator();
                    showError(formContext, `Error uploading file: ${error.message}`);
                    reject(error);
                }
            );
        };

        reader.onerror = function () {
            Xrm.Utility.closeProgressIndicator();
            showError(formContext, "Error reading the selected file.");
            reject(new Error("Error reading the selected file."));
        };

        reader.readAsDataURL(file); // Read the file as Base64
    });
}

// Get Attachment Config
function getAttachmentConfig(formContext) {
    const caseToken = formContext.getAttribute("ap_tsacasetoken").getValue();

    if (!caseToken) {
        showError(formContext, "CaseToken is required.");
        return Promise.reject(new Error("CaseToken is required."));
    }
    //Xrm.Utility.showProgressIndicator("Retrieving attachment config...");

    const parameters = {
        CaseToken: caseToken
    };

    const request = {
        CaseToken: parameters.CaseToken,
        getMetadata: function () {
            return {
                boundParameter: null,
                parameterTypes: {
                    "CaseToken": { typeName: "Edm.String", structuralProperty: 1 }
                },
                operationType: 0,
                operationName: "ap_GetAttachmentConfig"
            };
        }
    };

    return Xrm.WebApi.online.execute(request).then(
        function success(result) {
            if (result.ok) {
                return result.json().then(function (response) {
                    if (!response.IsError) {
                        //Xrm.Utility.closeProgressIndicator();
                        var configJson = response.GetAttachmentConfigResponse;
                        var configResponse = JSON.parse(configJson);
                        return configResponse;
                    } else {
                        //Xrm.Utility.closeProgressIndicator();
                        var error = JSON.parse(response.GetAttachmentConfigResponse);
                        showError(formContext, error.message);
                        return Promise.reject(new Error(error.message));
                    }
                });
            } else {
                //Xrm.Utility.closeProgressIndicator();
                return Promise.reject(new Error("Failed to retrieve attachment config."));
            }
        },
        function (error) {
            //Xrm.Utility.closeProgressIndicator();
            showError(formContext, error.message);
            return Promise.reject(new Error(error.message));
        }
    );
}
