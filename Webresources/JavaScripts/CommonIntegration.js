// Constant for standard alert dialog options
const STANDARD_ALERT_OPTIONS = { height: 120, width: 260 };

// Get Case
function getCase(formContext) {
    "use strict";
    const caseToken = formContext.getAttribute("ap_tsacasetoken").getValue();
    const entityId = formContext.data.entity.getId();

    if (!caseToken) {
        showError(formContext, "Case Token is required.");
        return;
    }
    Xrm.Utility.showProgressIndicator("Retrieving case details...");

    const parameters = {
        CaseToken: caseToken
    };

    const request = {
        entity: {
            id: entityId.replace(/[{}]/g, ""),
            entityType: "ap_tsanetcase"
        },
        CaseToken: parameters.CaseToken,
        getMetadata: function () {
            return {
                boundParameter: "entity",
                parameterTypes: {
                    "entity": {
                        "typeName": "mscrm.ap_tsanetcase",
                        "structuralProperty": 5
                    },
                    "CaseToken": {
                        "typeName": "Edm.String",
                        "structuralProperty": 1
                    }
                },
                operationType: 0,
                operationName: "ap_RefreshCase"
            };
        }
    };

    Xrm.WebApi.online.execute(request).then(
        function success(result) {
            if (result.ok) {
                result.json().then(function (response) {
                    if (!response.IsError) {
                        var formJson = response.CaseResponse;
                        var formResponse = JSON.parse(formJson);
                        saveToFormField("ap_formjson", formResponse, formContext);  // Save JSON
                        updateTsaCaseStatus(formContext, formResponse.status);    
                        Xrm.Utility.closeProgressIndicator();

                    } else {
                        Xrm.Utility.closeProgressIndicator();
                        var error = JSON.parse(response.CaseResponse);
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
// DEPRECATED
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
    // Save the record
    formContext.data.save().then(
        function success(result) {
            Xrm.Utility.showProgressIndicator("Sending case...");

            // Convert the submissionData object to a JSON string
            const submissionDataString = JSON.stringify(submissionData);
            const caseId = formContext.data.entity.getId().replace("{", "").replace("}", "");

            const parameters = {
                CaseDetails: submissionDataString,
                CaseID: caseId
            };

            // Custom action call
            const request = {
                CaseDetails: parameters.CaseDetails,
                CaseID: parameters.CaseID,
                getMetadata: function () {
                    return {
                        boundParameter: null, // No entity bound
                        parameterTypes: {
                            "CaseDetails": { typeName: "Edm.String", structuralProperty: 1 },
                            "CaseID": { typeName: "Edm.String", structuralProperty: 1 }
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
                                //fill data to form
                                formContext.getAttribute("ap_name").setValue(formResponse.id.toString());
                                formContext.getAttribute("ap_submittercasenumber").setValue(formResponse.submitterCaseNumber.toString());
                                formContext.getAttribute("ap_tsacasetoken").setValue(formResponse.token.toString());
                                saveToFormField("ap_formjson", formResponse, formContext);  // Save JSON
                                formContext.data.save();
                                Xrm.Utility.closeProgressIndicator();
                                showSuccess(formContext, "Successfully created!");
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
        },
        function(error) {
             showError(formContext, `Error saving updated case: ${error.message}`);
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
