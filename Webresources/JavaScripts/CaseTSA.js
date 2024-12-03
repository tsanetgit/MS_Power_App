// Start polling to ensure the input element is available
function onFormLoad(executionContext) {
    const formContext = executionContext.getFormContext();
    const formJsonField = formContext.getAttribute("ap_formjson").getValue();
    //waitForWebResourceElement('WebResource_casecreate', 'companyInput', () => setupCompanySearch(formContext));

    // Wait for the web resource element to load
    waitForWebResourceElement('WebResource_casecreate', 'dynamicFormContainer', () => {
        if (formJsonField) {
            // If `ap_formjson` contains data, parse it and build the read-only form            
            const formJsonData = JSON.parse(formJsonField);
            buildReadOnlyForm(formJsonData, formContext);
        } else {
            // Call your existing logic to display editable form
            setupCompanySearch(formContext);
        }
    });
}

// On form change
function onFormChange(executionContext) {
    const formContext = executionContext.getFormContext();
    const formJsonField = formContext.getAttribute("ap_formjson").getValue();

    if (formJsonField) {
        // If `ap_formjson` contains data, parse it and build the read-only form
        const formJsonData = JSON.parse(formJsonField);
        buildReadOnlyForm(formJsonData, formContext);
    } else {
        // Call your existing logic to display editable form
        setupCompanySearch(formContext);
    }
}

function buttonRefreshCase(formContext) {
    getCase(formContext);
}


// Setup the search functionality once the element is found
function setupCompanySearch(formContext) {
    const webResourceControl = parent.Xrm.Page.getControl("WebResource_casecreate");
    const webResourceContent = webResourceControl.getObject().contentDocument;

    // Show the company search element
    const companySearchElement = webResourceContent.getElementById("company-search");
    if (companySearchElement) {
        companySearchElement.style.display = "block";  // Show the company search section
    }

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
        showError(formContext, "Please enter a company name");
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
        const tags = company.tags.map(tag => tag.tag).join(", ");
        const departmentName = company.departmentName ? ` - ${company.departmentName}` : "";
        const tagsDisplay = tags ? ` [${tags}]` : "";
        option.text = `${company.companyName}${departmentName}${tagsDisplay}`;
        option.value = JSON.stringify({ companyName: company.companyName, companyId: company.companyId, departmentId: company.departmentId });
        selectList.appendChild(option);
    });

    // Add an event listener to trigger both selectCompany and getFormByCompany
    selectList.addEventListener("change", function () {
        const selectedValue = JSON.parse(selectList.value);
        const selectedCompanyName = selectedValue.companyName;
        if (selectedValue.companyId) {
            selectCompany(formContext, selectedValue.companyId, selectedCompanyName);
            getFormByCompany(selectedValue.companyId, formContext);
        }
    });

    resultsDiv.appendChild(selectList);
}

// Function to set selected company data into form fields
function selectCompany(formContext, companyId, companyName) {
    formContext.getAttribute("ap_companycode").setValue(companyId);
    formContext.getAttribute("ap_companyname").setValue(companyName);
}

// Function to dynamically create the form
function displayDynamicForm(formDetails, formContext) {
    const webResourceControl = parent.Xrm.Page.getControl("WebResource_casecreate");
    const webResourceContent = webResourceControl.getObject().contentDocument;

    const formContainer = webResourceContent.getElementById("dynamicFormContainer");
    formContainer.innerHTML = "";  // Clear existing form

    const form = document.createElement("form");
    form.className = "dynamic-form";

    // Add fields based on formDetails
    form.appendChild(createTextInput("Internal Case Number", formDetails.internalCaseNumber, "internalCaseNumber"));
    form.appendChild(createTextInput("Receiver Case Number", formDetails.receiverCaseNumber, "receiverCaseNumber"));
    form.appendChild(createTextInput("Problem Summary", formDetails.problemSummary, "problemSummary"));
    form.appendChild(createTextInput("Problem Description", formDetails.problemDescription, "problemDescription"));

    form.appendChild(createPrioritySelect("Priority", formDetails.priority, "priority"));
    form.appendChild(createHtmlField("Admin Note", formDetails.adminNote, "adminNote"));
    form.appendChild(createHtmlField("Escalation Instructions", formDetails.escalationInstructions, "escalationInstructions"));

    // Add internal note field as read-only
    const internalNote = formDetails.internalNotes && formDetails.internalNotes.length > 0 ? formDetails.internalNotes[0].note : "";
    form.appendChild(createReadOnlyTextArea("Internal Note", internalNote));

    // Custom fields
    const customerDataSections = groupBy(formDetails.customFields, "section");
    for (const section in customerDataSections) {
        const sectionGroup = document.createElement("div");
        sectionGroup.className = "form-section";
        sectionGroup.innerHTML = `<h3>${section}</h3>`;

        customerDataSections[section].sort((a, b) => a.displayOrder - b.displayOrder);
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
        if (validateForm(form)) {
            const submissionData = buildFormObject(formDetails);
            saveToFormField("ap_sentjson", submissionData, formContext);  // Save JSON
            postCase(submissionData, formContext);  // Send the object via existing unbound action
        } else {
            showError(formContext, "Please correct the data.");
        }
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

// Helper function to create text area inputs
function createTextArea(label, value, name) {
    const inputGroup = document.createElement("div");
    inputGroup.className = "input-group";

    const labelElement = document.createElement("label");
    labelElement.className = "form-label";
    labelElement.textContent = label;

    const textArea = document.createElement("textarea");
    textArea.value = value || "";
    textArea.name = name;
    textArea.className = "form-input";

    inputGroup.appendChild(labelElement);
    inputGroup.appendChild(textArea);

    return inputGroup;
}

// Helper function to create text area inputs
function createTextArea(label, value, name) {
    const inputGroup = document.createElement("div");
    inputGroup.className = "input-group";

    const labelElement = document.createElement("label");
    labelElement.className = "form-label";
    labelElement.textContent = label;

    const textArea = document.createElement("textarea");
    textArea.value = value || "";
    textArea.name = name;
    textArea.className = "form-input";

    inputGroup.appendChild(labelElement);
    inputGroup.appendChild(textArea);

    return inputGroup;
}

// Helper function for dynamically creating fields based on Type
function createFieldFromMetadata(field) {
    const fieldGroup = document.createElement("div");
    fieldGroup.className = "input-group";

    const labelElement = document.createElement("label");
    labelElement.className = "form-label";
    labelElement.textContent = field.label;

    let inputElement;
    switch (field.type.toLowerCase()) {
        case "integer":
            inputElement = document.createElement("input");
            inputElement.type = "number";
            inputElement.value = field.value || 0;
            inputElement.className = "form-input";
            break;
        case "email":
            inputElement = document.createElement("input");
            inputElement.type = "email";
            inputElement.value = field.value || "";
            inputElement.className = "form-input";
            break;
        case "phone":
            inputElement = document.createElement("input");
            inputElement.type = "tel";
            inputElement.pattern = "\\d+";  // Only digits
            inputElement.value = field.value || "";
            inputElement.className = "form-input";
            break;
        case "select":
            inputElement = document.createElement("select");
            inputElement.className = "form-input";
            (field.options || []).forEach(option => {
                const opt = document.createElement("option");
                opt.value = option;
                opt.textContent = option;
                inputElement.appendChild(opt);
            });
            break;
        case "tierselect":
            inputElement = createTierSelect(field.selections, field.value);
            inputElement.className = "form-input";
            break;
        default:
            inputElement = document.createElement("input");
            inputElement.type = "text";
            inputElement.value = field.value || "";
            inputElement.className = "form-input";
            break;
    }

    // Assign a unique name using the FieldId
    inputElement.name = `field_${field.fieldId}`;

    // Apply validation for "not_numeric"
    if (field.validationRules && field.validationRules.includes("not_numeric")) {
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

// Helper function for priority dropdown
function createPrioritySelect(label, value, name) {
    const inputGroup = document.createElement("div");
    inputGroup.className = "input-group";

    const labelElement = document.createElement("label");
    labelElement.className = "form-label";
    labelElement.textContent = label;

    const select = document.createElement("select");
    select.className = "form-input";
    select.name = name;

    ["LOW", "MEDIUM", "HIGH"].forEach(optionValue => {
        const option = document.createElement("option");
        option.value = optionValue;
        option.textContent = optionValue;
        if (optionValue.toLowerCase() === value.toLowerCase()) option.selected = true;
        select.appendChild(option);
    });

    inputGroup.appendChild(labelElement);
    inputGroup.appendChild(select);
    return inputGroup;
}

// Helper function for HTML read-only fields
function createHtmlField(label, value, name) {
    const inputGroup = document.createElement("div");
    inputGroup.className = "input-group";

    const labelElement = document.createElement("label");
    labelElement.className = "form-label";
    labelElement.textContent = label;

    const div = document.createElement("div");
    div.className = "readonly-html";
    div.innerHTML = value;  // Render HTML

    inputGroup.appendChild(labelElement);
    inputGroup.appendChild(div);
    return inputGroup;
}

// Helper function to validate the form
function validateForm(form) {   
    return form.checkValidity();  
}

// Helper function to build the form object, saving current form values
function buildFormObject(formDetails) {
    const formContext = parent.Xrm.Page.getControl("WebResource_casecreate").getObject().contentDocument;
    const cleanedObject = JSON.parse(JSON.stringify(formDetails));  // Deep clone

    // Update main fields with current values from the form
    cleanedObject.casePriority = formContext.querySelector('[name="priority"]').value;
    cleanedObject.internalCaseNumber = formContext.querySelector('[name="internalCaseNumber"]').value;
    cleanedObject.optionalRecieverInternalCaseNumber = formContext.querySelector('[name="receiverCaseNumber"]').value;
    cleanedObject.problemSummary = formContext.querySelector('[name="problemSummary"]').value;
    cleanedObject.problemDescription = formContext.querySelector('[name="problemDescription"]').value;

    // For customer fields, update the current values from the form inputs
    cleanedObject.customFields.forEach(data => {
        const fieldElement = formContext.querySelector(`[name="field_${data.fieldId}"]`);

        if (fieldElement) {
            if (fieldElement.tagName === "SELECT") {
                data.value = fieldElement.value;  // For select dropdowns
            } else {
                data.value = fieldElement.value;  // For text, number, email, etc.
            }
        }

        // Remove metadata and selections to keep the structure clean
        //delete data.selections;
    });

    return cleanedObject;
}

function buildReadOnlyForm(formJsonData, formContext) {
    const webResourceControl = parent.Xrm.Page.getControl("WebResource_casecreate");
    const webResourceContent = webResourceControl.getObject().contentDocument;

    const formContainer = webResourceContent.getElementById("dynamicFormContainer");
    formContainer.innerHTML = "";  // Clear existing form

    const form = document.createElement("form");
    form.className = "dynamic-form";

    // Add form fields as read-only, no company search or submit button
    form.appendChild(createReadOnlyTextField("Submitter Case Number", formJsonData.submitterCaseNumber));
    form.appendChild(createReadOnlyTextField("Receiver Case Number", formJsonData.receiverCaseNumber));
    form.appendChild(createReadOnlyTextField("Summary", formJsonData.summary));
    form.appendChild(createReadOnlyTextField("Description", formJsonData.description));

    // Add internal note field as read-only
    const internalNote = formJsonData.internalNotes && formJsonData.internalNotes.length > 0 ? formJsonData.internalNotes[0].note : "";
    form.appendChild(createReadOnlyTextArea("Internal Note", internalNote));

    form.appendChild(createReadOnlyTextField("Priority", formJsonData.priority));

    form.appendChild(createReadOnlyHtmlField("Escalation Instructions", formJsonData.escalationInstructions));
    form.appendChild(createReadOnlyHtmlField("Priority Note", formJsonData.priorityNote));

    // Add custom fields
    const customFields = groupBy(formJsonData.customFields, "section");
    for (const section in customFields) {
        const sectionGroup = document.createElement("div");
        sectionGroup.className = "form-section";
        sectionGroup.innerHTML = `<h3>${section}</h3>`;

        customFields[section].forEach(field => {
            sectionGroup.appendChild(createReadOnlyTextField(field.fieldName, field.value));
        });
        form.appendChild(sectionGroup);
    }

    formContainer.appendChild(form);
}

// Helper function for read-only text areas
function createReadOnlyTextArea(label, value) {
    const inputGroup = document.createElement("div");
    inputGroup.className = "input-group";

    const labelElement = document.createElement("label");
    labelElement.className = "form-label";
    labelElement.textContent = label;

    const textArea = document.createElement("textarea");
    textArea.value = value || "";
    textArea.className = "form-input";
    textArea.readOnly = true;

    inputGroup.appendChild(labelElement);
    inputGroup.appendChild(textArea);
    return inputGroup;
}

// Helper function for read-only text fields
function createReadOnlyTextField(label, value) {
    const inputGroup = document.createElement("div");
    inputGroup.className = "input-group";

    const labelElement = document.createElement("label");
    labelElement.className = "form-label";
    labelElement.textContent = label;

    const input = document.createElement("input");
    input.type = "text";
    input.value = value || "";
    input.className = "form-input";
    input.readOnly = true;

    inputGroup.appendChild(labelElement);
    inputGroup.appendChild(input);
    return inputGroup;
}

// Helper function for read-only HTML fields
function createReadOnlyHtmlField(label, value) {
    const inputGroup = document.createElement("div");
    inputGroup.className = "input-group";

    const labelElement = document.createElement("label");
    labelElement.className = "form-label";
    labelElement.textContent = label;

    const div = document.createElement("div");
    div.className = "readonly-html";
    div.innerHTML = value || "";  // Render HTML content

    inputGroup.appendChild(labelElement);
    inputGroup.appendChild(div);
    return inputGroup;
}
