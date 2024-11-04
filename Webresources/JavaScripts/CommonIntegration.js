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

                        // Process caseNotes array to patch each note
                        let caseNotes = formResponse.caseNotes;
                        let updatePromises = caseNotes.map(note => {
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

                        // Execute all PATCH requests and handle final actions
                        Promise.all(updatePromises)
                            .then(() => {
                                Xrm.Utility.closeProgressIndicator();
                                formContext.ui.setFormNotification("Successfully updated!", "INFO", "success");
                                // Refresh the notes subgrid
                                formContext.getControl("notesubgrid").refresh();
                                formContext.data.entity.save();
                            })
                            .catch(error => {
                                Xrm.Utility.closeProgressIndicator();
                                console.error(error.message);
                                showError(formContext, "Error updating case notes: " + error.message);
                            });

                    } else {
                        Xrm.Utility.closeProgressIndicator();
                        showError(formContext, response.GetCaseResponse);
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

// Helper function to map priority from string to option set value
function getPriorityValue(priority) {
    switch (priority.toLowerCase()) {
        case "low": return 3;
        case "medium": return 2;
        case "high": return 1;
        default: return null;
    }
}


// Get company
function getCompanyDetails(companyName) {
    return new Promise(function (resolve, reject) {
        Xrm.Utility.showProgressIndicator("Retrieving company details...");
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
                            Xrm.Utility.closeProgressIndicator();
                            resolve(companyDetails);
                        });
                } else {
                    Xrm.Utility.closeProgressIndicator();
                    reject(new Error("No result from company search"));
                }
            },
            function (error) {
                Xrm.Utility.closeProgressIndicator();
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
                        showError(formContext, response.ErrorMessage);
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
                        showError(formContext, response.PostCaseResponse);
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