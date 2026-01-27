"use strict";

// Global flag to prevent save loop
var isSaving = false;

// On form load function that initializes the form
function onFormLoad(executionContext) {
    const formContext = executionContext.getFormContext();
    const formJsonField = formContext.getAttribute("ap_formjson").getValue();
    //waitForWebResourceElement('WebResource_casecreate', 'companyInput', () => setupCompanySearch(formContext));
    statusWarning(executionContext);
    registerEventHandlers(executionContext);
    // Wait for the web resource element to load
    waitForWebResourceElement(formContext, 'WebResource_casecreate', 'dynamicFormContainer', () => {
        if (formJsonField) {
            initializeUploadNotificationMonitoring(formContext);
        } else if (formContext.ui.getFormType() === 1)  {
            // Call your existing logic to display editable form
            setupCompanySearch(formContext);
        }
    });
}

function registerEventHandlers(executionContext) {
    const formContext = executionContext.getFormContext();
    // Register onChange event for the statuscode field
    formContext.getAttribute("statuscode").addOnChange(function () {
        statusWarning(executionContext);
    });
    // Register onChange event for ap_formjson field
    const formJsonAttribute = formContext.getAttribute("ap_formjson");
    if (formJsonAttribute) {
        formJsonAttribute.addOnChange(onFormJsonChange);
    }
    // Register the onSave event handler
    formContext.data.entity.addOnSave(onSave);
}

function statusWarning(executionContext) {
    const formContext = executionContext.getFormContext();
    TSA.statusWarningLogic(formContext);
}

/**
 * OnChange handler for ap_formjson field
 * Triggers webresource refresh when JSON is updated
 */
function onFormJsonChange(executionContext) {
    const formContext = executionContext.getFormContext();
    const webResourceControl = formContext.getControl("WebResource_casecreate");

    const webResourceWindow = webResourceControl.getObject().contentWindow;

    if (webResourceWindow && webResourceWindow.refreshReadOnlyForm) {
        webResourceWindow.refreshReadOnlyForm();
    }
}

function buttonRefreshCase(formContext) {
    getCase(formContext);
}
function setupCompanySearch(formContext) {
    const webResourceControl = formContext.getControl("WebResource_casecreate");
    const webResourceContent = webResourceControl.getObject().contentDocument;

    // Show the company search element
    const companySearchElement = webResourceContent.getElementById("company-search");
    if (companySearchElement) {
        companySearchElement.style.display = "block";  // Show the company search section
    }

    const companyInput = webResourceContent.getElementById("companyInput");

    if (companyInput) {
        companyInput.addEventListener("input", function (event) {
            const inputValue = companyInput.value.trim();
            const wordCount = inputValue.length;

            if (wordCount >= 3) {
                searchCompany(formContext, inputValue);
            }
        });
    } else {
        showError(formContext, "companyInput element not found");
    }

    // footer with contact us link
    const footer = document.createElement("div");
    footer.className = "company-search-footer";
    footer.innerHTML = 'Not able to find a Member? <a href="mailto:connect_support@tsanet.org">Contact Us</a>';
    companySearchElement.appendChild(footer);

    // Make an initial API request to warm up the API
    getCompanyDetails("Cold API Request").then(function () {
        
    }).catch(function (error) {
        showError(formContext, "Error during initial API request: " + error.message);
    });
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
        showError(formContext, "Error retrieving company details: " + error.message);
    });
}

// Function to display and select a company
function displayCompanyResults(formContext, companies) {
    const webResourceControl = formContext.getControl("WebResource_casecreate");
    const webResourceContent = webResourceControl.getObject().contentDocument;

    const companySearchElement = webResourceContent.getElementById("company-search");
    const companyInput = webResourceContent.getElementById("companyInput");
    const resultsDiv = webResourceContent.getElementById("companyResults");
    resultsDiv.innerHTML = "";  // Clear previous results

    if (!companies.length) {
        resultsDiv.innerHTML = "<p>No companies found.</p>";
    } else {
        // Create a dropdown (select element)
        const selectList = document.createElement("select");
        selectList.size = companies.length + 1;  // Show all options

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
                // Check if a department is selected then use method for department
                if (selectedValue.departmentId != null) {
                    getFormByDepartment(selectedValue.departmentId, formContext);
                } else {
                    getFormByCompany(selectedValue.companyId, formContext);
                }
                // Hide the company search section
                companySearchElement.style.display = "none";  
            }
        });

        resultsDiv.appendChild(selectList);
        resultsDiv.style.width = companyInput.offsetWidth + "px";
        resultsDiv.style.top = companyInput.offsetTop + companyInput.offsetHeight + "px";
        resultsDiv.style.left = companyInput.offsetLeft + "px";
    }

    // Ensure the footer is always displayed below the results or the "No companies found" message
    const footer = webResourceContent.querySelector(".company-search-footer");
    if (footer) {
        resultsDiv.appendChild(footer);
    } else {
        // If the footer is not found, create and append it
        const newFooter = document.createElement("div");
        newFooter.className = "company-search-footer";
        newFooter.innerHTML = 'Not able to find a Member? <a href="mailto:connect_support@tsanet.org">Contact Us</a>';
        resultsDiv.appendChild(newFooter);
    }
}

// Function to set selected company data into form fields
function selectCompany(formContext, companyId, companyName) {
    formContext.getAttribute("ap_companycode").setValue(companyId);
    formContext.getAttribute("ap_companyname").setValue(companyName);
}

// Function that builds dynamic form based on company data
async function displayDynamicForm(formDetails, formContext) {
    const webResourceControl = formContext.getControl("WebResource_casecreate");
    const webResourceContent = webResourceControl.getObject().contentDocument;

    const formContainer = webResourceContent.getElementById("dynamicFormContainer");
    formContainer.innerHTML = "";  // Clear existing form

    // Store formDetails as a data attribute for later retrieval
    formContainer.dataset.formDetails = JSON.stringify(formDetails);

    const form = document.createElement("form");
    form.className = "dynamic-form";

    // Container for left and right sections
    const sectionsContainer = document.createElement("div");
    sectionsContainer.className = "sections-container";

    // Left section
    const leftSection = document.createElement("div");
    leftSection.className = "form-left-section";

    leftSection.appendChild(createTextInput("Company", formContext.getAttribute("ap_companyname").getValue(), "partnerName", true, false, false, 100));

    leftSection.appendChild(createHtmlField("Admin Note", formDetails.adminNote, "adminNote"));
    // Internal note field as read-only
    const internalNote = formDetails.internalNotes && formDetails.internalNotes.length > 0 ? formDetails.internalNotes[0].note : "";
    leftSection.appendChild(createHtmlField("Internal Note", internalNote, "internalNote"));

    // Retrieve mapping configurations and incident data if ap_caseid is present
    const apCase = formContext.getAttribute("ap_caseid");
    let incident = null;
    let mappingConfigs = [];
    let incidentMappedData = {};

    if (apCase) {
        const apCaseId = apCase.getValue();
        if (apCaseId && apCaseId[0] && apCaseId[0].id) {
            const caseId = apCaseId[0].id.replace(/[{}]/g, '');

            // Retrieve mapping configurations
            try {
                mappingConfigs = await retrieveFormMappingConfigurations();
            } catch (error) {
                showError(formContext, "Error retrieving mapping configurations: " + error.message);
            }

            // Build the select query with expands based on mappings
            const selectQuery = buildIncidentSelectQuery(mappingConfigs);

            try {
                incident = await Xrm.WebApi.retrieveRecord("incident", caseId, selectQuery);

                // Process mappings to get the mapped data from case
                incidentMappedData = processMappingsForForm(incident, mappingConfigs);
            } catch (error) {
                showError(formContext, "Error retrieving incident data: " + error.message);
            }
        }
    }

    // second section - use mapped data or formDetails as fallback
    leftSection.appendChild(createTextInput("Internal Case#",
        incidentMappedData["internalCaseNumber"] || formDetails.internalCaseNumber,
        "internalCaseNumber", false, true, true, 100));
    leftSection.appendChild(createPrioritySelect("Priority", formDetails.priority, "priority"));
    leftSection.appendChild(createTextInput("Summary",
        incidentMappedData["problemSummary"] || formDetails.problemSummary,
        "problemSummary", false, true, false, 300));
    leftSection.appendChild(createTextAreaInput("Description",
        incidentMappedData["problemDescription"] || formDetails.problemDescription,
        "problemDescription", false, true));

    // Right section
    const rightSection = document.createElement("div");
    rightSection.className = "form-right-section";

    // Custom fields
    const customerDataSections = groupBy(formDetails.customFields, "section");
    for (const section in customerDataSections) {
        const sectionGroup = document.createElement("div");
        sectionGroup.className = "form-section";

        // Map section names to user-friendly display names
        let displaySectionName = section;
        if (section.toUpperCase() === "CONTACT_SECTION") {
            displaySectionName = "Contact";
        } else if (section.toUpperCase() === "COMMON_CUSTOMER_SECTION") {
            displaySectionName = "Customer Details";
        } else if (section.toUpperCase() === "PROBLEM_SECTION") {
            displaySectionName = "Problem Details";
        }

        sectionGroup.innerHTML = `<h3>${displaySectionName}</h3>`;

        customerDataSections[section].sort((a, b) => a.displayOrder - b.displayOrder);
        customerDataSections[section].forEach(field => {
            // Check if there's a mapped value for this field
            const mappedValue = incidentMappedData[field.label] || incidentMappedData[`field_${field.fieldId}`];
            if (mappedValue !== undefined && mappedValue !== null && mappedValue !== "") {
                field.value = mappedValue;
            }

            // Create the field and check if a valid element was returned
            const fieldElement = createFieldFromMetadata(field);
            if (fieldElement) {
                sectionGroup.appendChild(fieldElement);
            }
        });

        if (section.toUpperCase() === "COMMON_CUSTOMER_SECTION") {
            rightSection.appendChild(sectionGroup);
        } else {
            leftSection.appendChild(sectionGroup);
        }
    }

    rightSection.appendChild(createHtmlField("Escalation Instructions", formDetails.escalationInstructions, "escalationInstructions"));

    sectionsContainer.appendChild(leftSection);
    sectionsContainer.appendChild(rightSection);
    form.appendChild(sectionsContainer);

    // Submit button
    const submitButton = document.createElement("button");
    submitButton.textContent = "Submit";
    submitButton.type = "submit";

    submitButton.addEventListener("click", function (event) {
        event.preventDefault();
        formContext.data.save();
    });

    form.appendChild(submitButton);
    formContainer.appendChild(form);
}

// Function to retrieve form mapping configurations (mappingtype = 120950002)
function retrieveFormMappingConfigurations() {
    return new Promise(function (resolve, reject) {
        var fetchXml = `
            <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
              <entity name="ap_casemapping">
                <attribute name="ap_casemappingid" />
                <attribute name="ap_name" />
                <attribute name="ap_sourcejsonpath" />
                <attribute name="ap_targetattribute" />
                <attribute name="ap_attributetype" />
                <order attribute="ap_name" descending="false" />
                <filter>
                   <condition attribute="ap_mappingtype" operator="eq" value="120950002" />
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

// Function to build incident select query with expands based on mappings
function buildIncidentSelectQuery(mappingConfigs) {
    const selectFields = new Set(["ticketnumber", "title", "description"]);
    const expandFields = {};

    mappingConfigs.forEach(function (mapping) {
        const sourcePath = mapping.ap_sourcejsonpath;
        if (sourcePath && sourcePath.includes('.')) {
            // Has a dot notation, need to expand
            const parts = sourcePath.split('.');
            const expandEntity = parts[0];
            const expandField = parts[1];

            if (!expandFields[expandEntity]) {
                expandFields[expandEntity] = new Set();
            }
            expandFields[expandEntity].add(expandField);
        } else if (sourcePath) {
            // Direct field on incident
            selectFields.add(sourcePath);
        }
    });

    // Build the query string
    let query = "?$select=" + Array.from(selectFields).join(",");

    // Add expands
    const expandParts = [];
    for (const [entity, fields] of Object.entries(expandFields)) {
        const fieldsList = Array.from(fields).join(",");
        expandParts.push(`${entity}($select=${fieldsList})`);
    }

    if (expandParts.length > 0) {
        query += "&$expand=" + expandParts.join(",");
    }

    return query;
}

// Function to process mappings and extract values from incident
function processMappingsForForm(incident, mappingConfigs) {
    const mappedData = {};

    mappingConfigs.forEach(function (mapping) {
        const sourcePath = mapping.ap_sourcejsonpath;
        const targetAttribute = mapping.ap_targetattribute;

        if (!sourcePath || !targetAttribute) {
            return;
        }

        let value = null;

        if (sourcePath.includes('.')) {
            // Handle dot notation (e.g., "primarycontactid.emailaddress1")
            const parts = sourcePath.split('.');
            const expandEntity = parts[0];
            const expandField = parts[1];

            if (incident[expandEntity] && incident[expandEntity][expandField]) {
                value = incident[expandEntity][expandField];
            }
        } else {
            // Direct field on incident
            if (incident[sourcePath]) {
                value = incident[sourcePath];
            }
        }

        // Only set the value if it's not null/undefined/empty
        if (value !== null && value !== undefined && value !== "") {
            mappedData[targetAttribute] = value;
        }
    });

    return mappedData;
}

// Helper function to create text inputs with two-column layout
function createTextInput(label, value, name, isReadOnly, isRequired, validateUrl, maxLength) {
    const inputGroup = document.createElement("div");
    inputGroup.className = "input-group";

    const labelElement = document.createElement("label");
    labelElement.className = "form-label";
    labelElement.textContent = label;
    inputGroup.appendChild(labelElement);

    const input = document.createElement("input");
    input.type = "text";
    input.value = value || "";
    input.name = name;
    input.className = "form-input";
    input.setAttribute("data-maxlength", maxLength);

    //field is read only
    if (isReadOnly) {
        input.readOnly = true;
        input.classList.add("readonly-input");
    }
    //field is required
    if (isRequired) {
        input.required = true;
    }
    //field shouldn't contain URLs
    if (validateUrl) {
        input.setAttribute("data-validate-url", "true");
    }

    // Wrap input with wrapper and optional asterisk
    const inputWithWrapper = wrapInputWithWrapper(input, isRequired);
    inputGroup.appendChild(inputWithWrapper);

    return inputGroup;
}

// Helper function to create text inputs with
function createTextAreaInput(label, value, name, isReadOnly, isRequired) {
    const inputGroup = document.createElement("div");
    inputGroup.className = "input-group";

    const labelElement = document.createElement("label");
    labelElement.className = "form-label";
    labelElement.textContent = label;
    inputGroup.appendChild(labelElement);

    const textArea = document.createElement("textarea");
    textArea.value = value || "";
    textArea.name = name;
    textArea.className = "form-input";
    textArea.rows = 5;

    if (isReadOnly) {
        textArea.readOnly = true;
        textArea.classList.add("readonly-input");
    }

    if (isRequired) {
        textArea.required = true;
    }

    // Wrap textarea with wrapper and optional asterisk
    const textAreaWithWrapper = wrapInputWithWrapper(textArea, isRequired);
    inputGroup.appendChild(textAreaWithWrapper);

    return inputGroup;
}

// Helper function for dynamically creating fields based on Type
function createFieldFromMetadata(field) {
    // Check if field type is "hidden" - if so, don't create a field on the form
    if (field.type && field.type.toLowerCase() === "hidden") {
        return null; // Return null to indicate no field should be added
    }

    const fieldGroup = document.createElement("div");
    fieldGroup.className = "input-group";

    const labelElement = document.createElement("label");
    labelElement.className = "form-label";
    labelElement.textContent = field.label;
    fieldGroup.appendChild(labelElement);

    let inputElement;
    switch (field.type.toLowerCase()) {
        case "integer":
            inputElement = document.createElement("input");
            inputElement.type = "number";
            inputElement.value = field.value || 0;
            inputElement.className = "form-input";
            inputElement.name = `field_${field.fieldId}`;
            break;
        case "email":
            inputElement = document.createElement("input");
            inputElement.type = "email";
            inputElement.value = field.value || "";
            inputElement.className = "form-input";
            inputElement.name = `field_${field.fieldId}`;
            break;
        case "phone":
            inputElement = document.createElement("input");
            inputElement.type = "tel";
            inputElement.pattern = "\\d+";  // Only digits
            inputElement.value = field.value || "";
            inputElement.className = "form-input";
            inputElement.name = `field_${field.fieldId}`;
            break;
        case "select":
            inputElement = document.createElement("select");
            inputElement.className = "form-input";
            const optionsArray = field.options ? field.options.split(/\r?\n/).filter(option => option.trim() !== "") : [];
            optionsArray.forEach(option => {
                const opt = document.createElement("option");
                opt.value = option;
                opt.textContent = option;
                inputElement.appendChild(opt);
            });
            inputElement.name = `field_${field.fieldId}`;
            break;
        case "tierselect":
            inputElement = createTierSelect(field.selections, field.value, field.required, field.fieldId);
            inputElement.className = "form-input";
            break;
        default:
            inputElement = document.createElement("input");
            inputElement.type = "text";
            inputElement.value = field.value || "";
            inputElement.className = "form-input";
            inputElement.name = `field_${field.fieldId}`;
            break;
    }

    // Apply validation for "not_numeric"
    if (field.validationRules && field.validationRules.includes("not_numeric")) {
        inputElement.pattern = "\\D*";  // Regex for non-numeric input
    }

    if (field.required) {
        inputElement.required = true;
    }

    // Wrap input element with wrapper and optional asterisk
    const inputWithWrapper = wrapInputWithWrapper(inputElement, field.required);
    fieldGroup.appendChild(inputWithWrapper);

    return fieldGroup;
}

// Function to wrap input with wrapper and add asterisk or placeholder
function wrapInputWithWrapper(inputElement, isRequired) {
    // Create input wrapper
    const inputWrapper = document.createElement("div");
    inputWrapper.className = "input-wrapper";

    // Create the asterisk span
    const asterisk = document.createElement("span");
    asterisk.className = "required-asterisk";

    if (isRequired) {
        asterisk.textContent = "*";
    } else {
        // Add a transparent placeholder
        asterisk.innerHTML = "&#8203;"; // Zero-width space character
    }

    inputWrapper.appendChild(asterisk);

    // Append input element to the wrapper
    inputWrapper.appendChild(inputElement);

    return inputWrapper;
}


// Helper function to create TIERSELECT dropdowns
function createTierSelect(options, selectedValue, isRequired, fieldId) {
    const container = document.createElement("div");
    container.className = "tier-select-container";

    function createSelectElement(optionList, level) {
        const selectElement = document.createElement("select");
        selectElement.className = `tier-select level-${level}`;
        selectElement.name = `field_${fieldId}-${level}`;
        selectElement.setAttribute("data-tierselect", "true"); // Add a unique attribute

        const defaultOption = document.createElement("option");
        defaultOption.value = "";
        defaultOption.textContent = "Select an option";
        selectElement.appendChild(defaultOption);

        optionList.forEach(option => {
            const opt = document.createElement("option");
            opt.value = option.value;
            opt.textContent = option.value;
            selectElement.appendChild(opt);
        });

        selectElement.addEventListener("change", function () {
            // Remove all subsequent select elements
            const nextLevel = level + 1;
            const nextSelectElements = container.querySelectorAll(`.tier-select.level-${nextLevel}`);
            nextSelectElements.forEach(el => el.remove());

            // If an option is selected and it has children, create the next select element
            const selectedOption = optionList.find(opt => opt.value === selectElement.value);
            if (selectedOption && selectedOption.children && selectedOption.children.length) {
                const nextSelectElement = createSelectElement(selectedOption.children, nextLevel);
                container.appendChild(nextSelectElement);
            }
        });

        // Add space between selection boxes
        selectElement.style.marginBottom = "10px";

        // Set the required attribute if needed
        if (isRequired) {
            selectElement.required = true;
        }

        return selectElement;
    }

    // Initialize the first select element
    const initialSelectElement = createSelectElement(options, 1);
    container.appendChild(initialSelectElement);

    // Set the selected value if provided
    if (selectedValue) {
        let currentOptions = options;
        let currentSelectElement = initialSelectElement;
        const selectedValues = selectedValue.split(" > ");

        selectedValues.forEach((value, index) => {
            currentSelectElement.value = value;
            const selectedOption = currentOptions.find(opt => opt.value === value);
            if (selectedOption && selectedOption.children && selectedOption.children.length) {
                const nextSelectElement = createSelectElement(selectedOption.children, index + 2);
                container.appendChild(nextSelectElement);
                currentSelectElement = nextSelectElement;
                currentOptions = selectedOption.children;
            } else {
                // Set the name attribute for the current select element if it is the lowest level
                currentSelectElement.name = `field_${fieldId}`;
            }
        });
    }

    return container;
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

// Helper function for priority dropdown
function createPrioritySelect(label, value, name) {
    const inputGroup = document.createElement("div");
    inputGroup.className = "input-group";

    const labelElement = document.createElement("label");
    labelElement.className = "form-label";
    labelElement.textContent = label;
    inputGroup.appendChild(labelElement);

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

    // Wrap select with wrapper and optional asterisk
    const selectWithWrapper = wrapInputWithWrapper(select, true);
    inputGroup.appendChild(selectWithWrapper);

    return inputGroup;
}

// Helper function for HTML read-only fields
function createHtmlField(label, value, name) {
    const inputGroup = document.createElement("div");
    inputGroup.className = "input-group";

    if (label) {
        const labelElement = document.createElement("label");
        labelElement.className = "form-label";
        labelElement.textContent = label;
        inputGroup.appendChild(labelElement);
    }

    const field = document.createElement("div");
    field.className = "readonly-html";
    field.innerHTML = value;  // Render HTML

    const fieldWithWrapper = wrapInputWithWrapper(field, false);
    inputGroup.appendChild(fieldWithWrapper);

    return inputGroup;
}

// Helper function to validate the form
function validateForm(form, formContext) {
    // Standard form validation
    let isFormValid = form.checkValidity();

    Array.from(form.elements).forEach(el => {
        el.classList.remove('invalid');

        if (!el.checkValidity()) {
            el.classList.add('invalid');
        }

        //validate max length for text inputs
        const maxLen = el.dataset.maxlength ? parseInt(el.dataset.maxlength, 10) : null;
        if (el.type === "text" && maxLen && el.value.length > maxLen) {
            el.classList.add('invalid');
            isFormValid = false;
        }
    });

    //show error if form is invalid
    if (!isFormValid) {
        TSA.showError(formContext, "Please correct the data.");
        return isFormValid;
    }

    // Get all text inputs that should be validated for URLs
    const textInputs = form.querySelectorAll('input[type="text"][data-validate-url="true"]');

    textInputs.forEach(input => {
        const value = input.value.toLowerCase();
        if (value.includes("http://") || value.includes("https://")) {
            TSA.showError(formContext, "URLs are not recommended in some fields. Please review your entries.");
            isFormValid = false;
            input.classList.add('invalid');
        }
        else {
            input.classList.remove('invalid');
        }
    });

    return isFormValid;
}

// Helper function to build the form object, saving current form values
async function buildFormObject(formDetails, formContext) {
    const webresourceContent = formContext.getControl("WebResource_casecreate").getObject().contentDocument;
    const cleanedObject = JSON.parse(JSON.stringify(formDetails));  // Deep clone

    // Update main fields with current values from the form
    cleanedObject.priority = webresourceContent.querySelector('[name="priority"]').value;
    cleanedObject.internalCaseNumber = webresourceContent.querySelector('[name="internalCaseNumber"]').value;
    cleanedObject.problemSummary = webresourceContent.querySelector('[name="problemSummary"]').value;
    cleanedObject.problemDescription = webresourceContent.querySelector('[name="problemDescription"]').value;

    // Get current user details from Dataverse
    const userDetails = await TSA.getCurrentUserDetails();

    // Add submitterContactDetails object with user information
    cleanedObject.submitterContactDetails = {
        name: userDetails.name,
        email: userDetails.email,
        phone: userDetails.phone
    };

    // For customer fields, update the current values from the form inputs
    cleanedObject.customFields.forEach(data => {
        const fieldElement = webresourceContent.querySelector(`[name*="field_${data.fieldId}"]`);

        if (fieldElement) {
            if (fieldElement.tagName === "SELECT" && fieldElement.getAttribute("data-tierselect") === "true") {
                // Get the value from the most child optionset
                const tierSelectElements = webresourceContent.querySelectorAll(`[name^="field_${data.fieldId}"]`);
                data.value = Array.from(tierSelectElements).map(el => el.value).join(' : ');

            } else if (fieldElement.tagName === "SELECT") {
                data.value = fieldElement.value;  // For select dropdowns
            } else {
                data.value = fieldElement.value;  // For text, number, email, etc.
            }
        } else {
            showError(formContext, `Field element not found for field ID: ${data.fieldId}`);
        }

        // Remove metadata and selections to keep the structure clean
        //delete data.selections;
    });

    return cleanedObject;
}

// OnSave event - main handler submit the case form
function onSave(executionContext) {
    const formContext = executionContext.getFormContext();
    const eventArgs = executionContext.getEventArgs();

    // Only run this logic if the form hasn't been submitted yet
    if (formContext.ui.getFormType() === 1 && !isSaving) {
        // Prevent the save until validation and data processing complete
        eventArgs.preventDefault();

        const webResourceControl = formContext.getControl("WebResource_casecreate");
        if (!webResourceControl) return;

        const webResourceContent = webResourceControl.getObject().contentDocument;
        const formContainer = webResourceContent.getElementById("dynamicFormContainer");
        const form = formContainer.querySelector(".dynamic-form");

        if (!form) return;

        if (validateForm(form, formContext)) {
            // Set the flag to prevent re-entry
            isSaving = true;

            // Retrieve formDetails from the web resource's data attribute
            const formDetails = getFormDetailsFromWebResource(webResourceContent);

            if (!formDetails) {
                showError(formContext, "Form details not available");
                // Reset flag
                isSaving = false;
                return;
            }

            buildFormObject(formDetails, formContext).then(submissionData => {
                saveToFormField("ap_sentjson", submissionData, formContext);
                postCase(submissionData, formContext);
            }).catch(error => {
                showError(formContext, "Error building form data: " + error.message);
                // Reset flag
                isSaving = false;
            });
        }
    }
}

// Helper function to retrieve formDetails from the web resource
function getFormDetailsFromWebResource(webResourceContent) {
    // Check if formDetails is stored as a data attribute on the form container
    const formContainer = webResourceContent.getElementById("dynamicFormContainer");
    if (formContainer && formContainer.dataset.formDetails) {
        return JSON.parse(formContainer.dataset.formDetails);
    }
    return null;
}