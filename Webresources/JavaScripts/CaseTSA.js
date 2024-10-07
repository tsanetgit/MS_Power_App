// Start polling to ensure the input element is available
function onFormLoad(executionContext) {
    const formContext = executionContext.getFormContext();
    waitForWebResourceElement('WebResource_casecreate', 'companyInput', () => setupCompanySearch(formContext));
}

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
                        displayDynamicForm(formDetails);
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

// Function to dynamically create the form
function displayDynamicForm(formDetails) {
    const webResourceControl = parent.Xrm.Page.getControl("WebResource_casecreate");
    const webResourceContent = webResourceControl.getObject().contentDocument;

    const formContainer = webResourceContent.getElementById("dynamicFormContainer");
    formContainer.innerHTML = "";  // Clear existing form if any

    // Create the form
    const form = document.createElement("form");
    form.className = "dynamic-form";

    // Add fields based on formDetails
    form.appendChild(createTextInput("Internal Case Number", formDetails.InternalCaseNumber, "internalCaseNumber"));
    form.appendChild(createTextInput("Receiver Case Number", formDetails.OptionalRecieverInternalCaseNumber, "receiverCaseNumber"));
    form.appendChild(createTextInput("Problem Summary", formDetails.ProblemSummary, "problemSummary"));
    form.appendChild(createTextInput("Problem Description", formDetails.ProblemDescription, "problemDescription"));
    form.appendChild(createTextInput("Priority", formDetails.CasePriority, "priority"));
    form.appendChild(createTextInput("Admin Note", formDetails.ReadonlyAdminNote, "adminNote", true));
    form.appendChild(createTextInput("Escalation Instructions", formDetails.ReadonlyEscalationInstructions, "escalationInstructions", true));

    // CustomerData Fields
    const customerDataSections = groupBy(formDetails.CustomerData, "Section");

    for (const section in customerDataSections) {
        const sectionGroup = document.createElement("div");
        sectionGroup.className = "form-section";
        sectionGroup.innerHTML = `<h3>${section}</h3>`;

        customerDataSections[section].sort((a, b) => a.FieldMetadata.DisplayOrder - b.FieldMetadata.DisplayOrder);

        customerDataSections[section].forEach(field => {
            sectionGroup.appendChild(createFieldFromMetadata(field));
        });

        form.appendChild(sectionGroup);
    }

    // Submit button
    const submitButton = document.createElement("button");
    submitButton.textContent = "Submit";
    submitButton.type = "submit";

    submitButton.addEventListener("click", function (event) {
        event.preventDefault();
        saveFormData(formDetails);
    });

    form.appendChild(submitButton);
    formContainer.appendChild(form);
}

// Helper function to create text inputs with two-column layout
function createTextInput(label, value, name, isReadOnly = false) {
    const inputGroup = document.createElement("div");
    inputGroup.className = "input-group";

    const labelElement = document.createElement("label");
    labelElement.className = "form-label";
    labelElement.textContent = label;

    const input = document.createElement("input");
    input.type = "text";
    input.value = value || "";
    input.name = name;
    input.className = "form-input";
    if (isReadOnly) {
        input.readOnly = true;
        input.classList.add("readonly-input");
    }

    inputGroup.appendChild(labelElement);
    inputGroup.appendChild(input);

    return inputGroup;
}

// Helper function to create fields based on metadata
function createFieldFromMetadata(field) {
    const fieldGroup = document.createElement("div");
    fieldGroup.className = "input-group";  // Use same class for two-column layout

    const labelElement = document.createElement("label");
    labelElement.className = "form-label";  // Style label
    labelElement.textContent = field.FieldMetadata.Label;

    let inputElement;
    switch (field.FieldMetadata.Type) {
        case "STRING":
            inputElement = document.createElement("input");
            inputElement.type = "text";
            inputElement.value = field.Value || "";
            inputElement.className = "form-input";  // Style input
            break;
        case "INT":
            inputElement = document.createElement("input");
            inputElement.type = "number";
            inputElement.value = field.Value || 0;
            inputElement.className = "form-input";  // Style input
            break;
        case "EMAIL":
            inputElement = document.createElement("input");
            inputElement.type = "email";
            inputElement.value = field.Value || "";
            inputElement.className = "form-input";  // Style input
            break;
        case "PHONE":
            inputElement = document.createElement("input");
            inputElement.type = "tel";
            inputElement.value = field.Value || "";
            inputElement.className = "form-input";  // Style input
            break;
        case "TIERSELECT":
            inputElement = createTierSelect(field.FieldMetadata.Options, field.Value);
            inputElement.className = "form-input";  // Style select
            break;
        default:
            inputElement = document.createElement("input");
            inputElement.type = "text";
            inputElement.value = field.Value || "";
            inputElement.className = "form-input";  // Style input
            break;
    }

    // Add validation, requirement, and custom styles
    if (field.FieldMetadata.Required) {
        inputElement.required = true;
    }
    if (field.FieldMetadata.ValidationRules && field.FieldMetadata.ValidationRules.includes("not_numeric")) {
        inputElement.pattern = "\\D*";  // Regex for non-numeric input
    }

    fieldGroup.appendChild(labelElement);
    fieldGroup.appendChild(inputElement);
    return fieldGroup;
}

// Helper function to create TIERSELECT dropdowns
function createTierSelect(options, selectedValue) {
    const selectElement = document.createElement("select");

    function createOptions(optionList, parentElement) {
        optionList.forEach(option => {
            const opt = document.createElement("option");
            opt.value = option.value;
            opt.textContent = option.value;
            parentElement.appendChild(opt);

            if (option.children && option.children.length) {
                createOptions(option.children, parentElement);
            }
        });
    }

    createOptions(options, selectElement);
    selectElement.value = selectedValue || "";
    return selectElement;
}

// Helper function to group data by section
function groupBy(arr, key) {
    return arr.reduce((group, item) => {
        const section = item[key];
        group[section] = group[section] || [];
        group[section].push(item);
        return group;
    }, {});
}

// Function to handle form submission and save data
function saveFormData(formDetails) {
    // Here you would gather data from the form and update the formDetails object
    console.log("Form data would be saved:", formDetails);
}