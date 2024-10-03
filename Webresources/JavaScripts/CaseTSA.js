function waitForElementsToExist(elementIds, callback, options) {
    options = Object.assign({
        checkFrequency: 500, // check for elements every 500 ms
        timeout: null, // after checking for X amount of ms, stop checking
    }, options);

    // poll every X amount of ms for all DOM nodes
    let intervalHandle = setInterval(() => {
        let doElementsExist = true;
        for (let elementId of elementIds) {
            let element = window.top.document.getElementById(elementId);
            if (!element) {
                // if element does not exist, set doElementsExist to false and stop the loop
                doElementsExist = false;
                break;
            }
        }

        // if all elements exist, stop polling and invoke the callback function 
        if (doElementsExist) {
            clearInterval(intervalHandle);
            if (callback) {
                callback();
            }
        }
    }, options.checkFrequency);
    if (options.timeout != null) {
        setTimeout(() => clearInterval(intervalHandle), options.timeout);
    }
}

function waitForWebResourceElement(webResourceName, elementId, callback, checkFrequency = 500, timeout = 5000) {
    const startTime = Date.now();
    const interval = setInterval(function () {
        try {
            const webResourceControl = parent.Xrm.Page.getControl(webResourceName);
            if (webResourceControl) {
                const webResourceContent = webResourceControl.getObject().contentDocument;
                console.log("Web resource content found");

                const element = webResourceContent.getElementById(elementId);  // Directly access the element
                if (element) {
                    clearInterval(interval);
                    console.log("Element found! Executing callback.");
                    callback();
                } else {
                    console.log("Element not found yet");
                }
            } else {
                console.log("Web resource control not found");
            }
        } catch (e) {
            console.error("An error occurred: " + e.message);
        }

        if (Date.now() - startTime > timeout) {
            clearInterval(interval);
            console.error("Element not found within the timeout period.");
        }
    }, checkFrequency);
}

// Setup the search functionality once the element is found
function setupCompanySearch(formContext) {
    const webResourceControl = parent.Xrm.Page.getControl("WebResource_casecreate");
    const webResourceContent = webResourceControl.getObject().contentDocument;

    const companyInput = webResourceContent.getElementById("companyInput");
    const searchButton = webResourceContent.getElementById("searchCompanyButton");

    if (companyInput) {
        console.log("Setting up company search...");

        companyInput.addEventListener("keypress", function (event) {
            if (event.key === "Enter") {
                event.preventDefault();
                searchCompany(formContext, companyInput.value);
            }
        });

        searchButton.addEventListener("click", function () {
            searchCompany(formContext, companyInput.value);
        });
    } else {
        console.error("companyInput element not found during setup!");
    }
}
// Function to trigger search
function searchCompany(formContext, companyName) {
    if (!companyName) {
        alert("Please enter a company name");
        return;
    }

    getCompanyDetails(companyName).then(function (companies) {
        displayCompanyResults(formContext, companies);
    }).catch(function (error) {
        console.error("Error retrieving company details: " + error.message);
    });
}

// Function to display and select a company
function displayCompanyResults(formContext, companies) {
    const webResourceControl = parent.Xrm.Page.getControl("WebResource_casecreate");
    const webResourceContent = webResourceControl.getObject().contentDocument;

    const resultsDiv = webResourceContent.getElementById("companyResults");
    resultsDiv.innerHTML = "";  // Clear previous results

    if (!companies.length) {
        resultsDiv.innerHTML = "<p>No companies found.</p>";
        return;
    }

    // Create a dropdown (select element)
    const selectList = document.createElement("select");
    const defaultOption = document.createElement("option");
    defaultOption.text = "--Select--";
    selectList.appendChild(defaultOption);

    companies.forEach(function (company) {
        const option = document.createElement("option");
        option.text = company.name;
        option.value = company.id;  // Use company.id which contains the actual ID
        selectList.appendChild(option);
    });

    // Add an event listener to trigger both selectCompany and getFormByCompany
    selectList.addEventListener("change", function () {
        const selectedCompanyId = parseInt(selectList.value);
        const selectedCompanyName = selectList.options[selectList.selectedIndex].text;
        if (selectedCompanyId) {
            selectCompany(formContext, selectedCompanyId, selectedCompanyName);
            getFormByCompany(selectedCompanyId);
        }
    });

    resultsDiv.appendChild(selectList);
}

// getFormByCompany function
function getFormByCompany(companyId) {
    const parameters = {
        CompanyId: companyId
    };

    // Custom action call
    const request = {
        CompanyId: parameters.CompanyId, // Pass the companyId parameter
        getMetadata: function () {
            return {
                boundParameter: null, // No entity bound
                parameterTypes: {
                    "CompanyId": { typeName: "Edm.Int32", structuralProperty: 1 } // Integer type for the CompanyId
                },
                operationType: 0, // 0 for actions, 1 for functions
                operationName: "ap_GetFormByCompany" // The name of your custom action in Dynamics 365
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

// Function to set selected company data into form fields
function selectCompany(formContext, companyId, companyName) {
    formContext.getAttribute("ap_companycode").setValue(companyId);
    formContext.getAttribute("ap_companyname").setValue(companyName);
}

// Get company name
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

// Start polling to ensure the input element is available
function onFormLoad(executionContext) {
    const formContext = executionContext.getFormContext();
    waitForWebResourceElement('WebResource_casecreate', 'companyInput', () => setupCompanySearch(formContext));
}
