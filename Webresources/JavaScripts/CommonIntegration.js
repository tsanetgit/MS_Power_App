function getCompanyDetails(companyName) {
    var parameters = {};
    parameters.CompanyName = companyName;

    var request = {
        getcompanybynameRequest: {
            entity: parameters,

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
        }
    };

    Xrm.WebApi.online.execute(request).then(
        function success(result) {
            if (result.ok) {
                result.json().then(function (response) {
                    var companyDetailsJson = response.CompanyDetails;
                    var companyDetails = JSON.parse(companyDetailsJson);

                    companyDetails.forEach(function (company) {
                        console.log("ID: " + company.id);
                        console.log("Name: " + company.name);
                        console.log("Description: " + company.description);
                    });
                });
            }
        },
        function (error) {
            console.log(error.message);
        }
    );
}

function getFormByCompany(companyId) {
    var parameters = {};
    parameters.CompanyId = companyId;

    var request = {
        getformbycompanyRequest: {
            entity: parameters,

            getMetadata: function () {
                return {
                    boundParameter: null,
                    parameterTypes: {
                        "CompanyId": { typeName: "Edm.String", structuralProperty: 1 }
                    },
                    operationType: 0,
                    operationName: "ap_GetFormByCompany" 
                };
            }
        }
    };

    Xrm.WebApi.online.execute(request).then(
        function success(result) {
            if (result.ok) {
                result.json().then(function (response) {
                    var formJson = response.FormDetails;
                    var formDetails = JSON.parse(formJson);

                    console.log(formDetails);
                });
            }
        },
        function (error) {
            console.log(error.message);
        }
    );
}

function postCase(caseDetails) {
    var parameters = {};
    parameters.CaseDetails = caseDetails;

    var request = {
        postcaseRequest: {
            entity: parameters,

            getMetadata: function () {
                return {
                    boundParameter: null,
                    parameterTypes: {
                        "CaseDetails": { typeName: "Edm.String", structuralProperty: 1 } 
                    },
                    operationType: 0,
                    operationName: "ap_PostCase"
                };
            }
        }
    };

    Xrm.WebApi.online.execute(request).then(
        function success(result) {
            if (result.ok) {
                result.json().then(function (response) {
                    console.log("Post Case Response: ", response.PostCaseResponse);
                });
            }
        },
        function (error) {
            console.log(error.message);
        }
    );
}


