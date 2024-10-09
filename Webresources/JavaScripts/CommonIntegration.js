// Get company
function getCompanyDetails(companyName) {
    return new Promise(function (resolve, reject) {
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
                            resolve(companyDetails);
                        });
                } else {
                    reject(new Error("No result from company search"));
                }
            },
            function (error) {
                reject(error);
            }
        );
    });
}

// getFormByCompany function
function getFormByCompany(companyId, formContext) {
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
                        displayDynamicForm(formDetails, formContext);
                    }
                    else {
                        alert(response.ErrorMessage);
                    }
                });
            }
        },
        function (error) {
            console.error(error.message); // Log any errors in the console
        }
    );
}

function postCase(submissionData, formContext) {
    // Convert the submissionData object to a JSON string
    const submissionDataString = JSON.stringify(submissionData);

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
                        console.log(formJson);
                    }
                    else {
                        alert(response.ErrorMessage);
                    }
                });
            }
        },
        function (error) {
            console.error(error.message);
            alert(error.message);
        }
    );
}