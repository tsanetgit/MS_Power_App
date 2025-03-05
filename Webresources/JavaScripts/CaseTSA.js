// Start polling to ensure the input element is available
function onFormLoad(executionContext) {
    const formContext = executionContext.getFormContext();
    const formJsonField = formContext.getAttribute("ap_formjson").getValue();
    //waitForWebResourceElement('WebResource_casecreate', 'companyInput', () => setupCompanySearch(formContext));
    statusWarning(executionContext);
    // Wait for the web resource element to load
    waitForWebResourceElement('WebResource_casecreate', 'dynamicFormContainer', () => {
        if (formJsonField) {
            // If `ap_formjson` contains data, parse it and build the read-only form            
            const formJsonData = JSON.parse(formJsonField);
            buildReadOnlyForm(formJsonData, formContext);
            detectTabChange(executionContext);
        } else {
            // Call your existing logic to display editable form
            setupCompanySearch(formContext);
        }
    });
}

function detectTabChange(executionContext) {

    setInterval(function () {
        // Check if dynamicFormContainer is empty and refresh if needed
        const webResourceControl = parent.Xrm.Page.getControl("WebResource_casecreate");
        const webResourceContent = webResourceControl.getObject().contentDocument;
        const formContainer = webResourceContent.getElementById("dynamicFormContainer");

        if (formContainer && formContainer.innerHTML.trim() === "") {
            onFormChange(executionContext);
        }

    }, 2000);
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

function statusWarning(executionContext) {
    const formContext = executionContext.getFormContext();
    const statusCode = formContext.getAttribute("statuscode").getValue();
    const direction = formContext.getAttribute("ap_direction").getValue();

    // Show warning if statuscode is 1
    if (statusCode === 1 && direction === 0) {
        showWarningMessage(formContext, "Response needed (Accept, request information, or Reject)​");
    }
    else {
        clearWarningMessage(formContext);
    }
}

function buttonRefreshCase(formContext) {
    getCase(formContext);
}
function setupCompanySearch(formContext) {
    const webResourceControl = parent.Xrm.Page.getControl("WebResource_casecreate");
    const webResourceContent = webResourceControl.getObject().contentDocument;

    // Show the company search element
    const companySearchElement = webResourceContent.getElementById("company-search");
    if (companySearchElement) {
        companySearchElement.style.display = "block";  // Show the company search section
    }

    const companyInput = webResourceContent.getElementById("companyInput");

    if (companyInput) {
        console.log("Setting up company search...");

        companyInput.addEventListener("input", function (event) {
            const inputValue = companyInput.value.trim();
            const wordCount = inputValue.length;

            if (wordCount >= 3) {
                searchCompany(formContext, inputValue);
            }
        });
    } else {
        console.error("companyInput element not found during setup!");
    }

    // footer with contact us link
    const footer = document.createElement("div");
    footer.className = "company-search-footer";
    footer.innerHTML = 'Not able to find a Member? <a href="mailto:connect_support@tsanet.org">Contact Us</a>';
    companySearchElement.appendChild(footer);

    // Make an initial API request to warm up the API
    getCompanyDetails("Cold API Request").then(function () {
        console.log("Initial API request to warm up the API completed.");
    }).catch(function (error) {
        console.error("Error during initial API request: " + error.message);
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
        console.error("Error retrieving company details: " + error.message);
    });
}

// Function to display and select a company
function displayCompanyResults(formContext, companies) {
    const webResourceControl = parent.Xrm.Page.getControl("WebResource_casecreate");
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
                getFormByCompany(selectedValue.companyId, formContext);
                companySearchElement.style.display = "none";  // Hide the company search section
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
    const webResourceControl = parent.Xrm.Page.getControl("WebResource_casecreate");
    const webResourceContent = webResourceControl.getObject().contentDocument;

    const formContainer = webResourceContent.getElementById("dynamicFormContainer");
    formContainer.innerHTML = "";  // Clear existing form

    const form = document.createElement("form");
    form.className = "dynamic-form";

    // Container for left and right sections
    const sectionsContainer = document.createElement("div");
    sectionsContainer.className = "sections-container";

    // Left section
    const leftSection = document.createElement("div");
    leftSection.className = "form-left-section";

    leftSection.appendChild(createTextInput("Company", formContext.getAttribute("ap_companyname").getValue(), "partnerName", true, false));

    leftSection.appendChild(createHtmlField("Admin Note", formDetails.adminNote, "adminNote"));
    // Add internal note field as read-only
    const internalNote = formDetails.internalNotes && formDetails.internalNotes.length > 0 ? formDetails.internalNotes[0].note : "";
    leftSection.appendChild(createHtmlField("Internal Note", internalNote, "internalNote"));

    // Prefill fields if ap_caseid is present
    const apCase = formContext.getAttribute("ap_caseid");
    let incident = null;
    if (apCase) {
        const apCaseId = apCase.getValue();
        if (apCaseId && apCaseId[0] && apCaseId[0].id) {
            const caseId = apCaseId[0].id.replace(/[{}]/g, '');
            incident = await getIncidentData(caseId);
        }
    }

    // second section
    leftSection.appendChild(createTextInput("Internal Case#", incident ? incident.ticketnumber : formDetails.internalCaseNumber, "internalCaseNumber", false, true));
    leftSection.appendChild(createPrioritySelect("Priority", formDetails.priority, "priority"));
    leftSection.appendChild(createTextInput("Summary", incident ? incident.title : formDetails.problemSummary, "problemSummary", false, true));
    leftSection.appendChild(createTextAreaInput("Description", incident ? incident.description : formDetails.problemDescription, "problemDescription", false, true));

    // Right section
    const rightSection = document.createElement("div");
    rightSection.className = "form-right-section";

    // Custom fields
    const customerDataSections = groupBy(formDetails.customFields, "section");
    for (const section in customerDataSections) {
        const sectionGroup = document.createElement("div");
        sectionGroup.className = "form-section";
        sectionGroup.innerHTML = `<h3>${section}</h3>`;

        customerDataSections[section].sort((a, b) => a.displayOrder - b.displayOrder);
        customerDataSections[section].forEach(field => {
            if (section.toUpperCase() === "COMMON_CUSTOMER_SECTION" && incident) {
                if (field.label.includes("Customer Email") && incident.primarycontactid?.emailaddress1) {
                    field.value = incident.primarycontactid.emailaddress1;
                }
                if (field.label.includes("Customer Phone") && incident.primarycontactid?.mobilephone) {
                    field.value = incident.primarycontactid.mobilephone;
                }
                if (field.label.includes("Customer Name") && incident.primarycontactid?.fullname) {
                    field.value = incident.primarycontactid.fullname;
                }
                if (field.label.includes("Customer Company") && (incident.customerid_account?.name || incident.customerid_contact?.fullname)) {
                    field.value = incident.customerid_account?.name || incident.customerid_contact?.fullname;
                }
            }
            sectionGroup.appendChild(createFieldFromMetadata(field));
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

// Helper function to retrieve incident data
async function getIncidentData(caseId) {
    try {
        const result = await Xrm.WebApi.retrieveRecord("incident", caseId, "?$select=ticketnumber,title,description&$expand=primarycontactid($select=emailaddress1,mobilephone,fullname),customerid_account($select=name),customerid_contact($select=fullname)");
        return result;
    } catch (error) {
        console.error("Error retrieving incident data: ", error);
        return null;
    }
}

// Helper function to create text inputs with two-column layout
function createTextInput(label, value, name, isReadOnly, isRequired) {
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

    if (isReadOnly) {
        input.readOnly = true;
        input.classList.add("readonly-input");
    }

    if (isRequired) {
        input.required = true;
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

    // Assign a unique name using the FieldId
    console.log(`Created field element with name: field_${field.fieldId}`);

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

    console.log(`Appended field element with name: field_${field.fieldId} to the DOM`);

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
                console.log(`Appended child select element for level ${nextLevel}`);
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
    console.log("Appended initial select element");

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
                console.log(`Set selected value for level ${index + 1}: ${value}`);
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
function validateForm(form) {   
    return form.checkValidity();  
}

// Helper function to build the form object, saving current form values
function buildFormObject(formDetails) {
    const formContext = parent.Xrm.Page.getControl("WebResource_casecreate").getObject().contentDocument;
    const cleanedObject = JSON.parse(JSON.stringify(formDetails));  // Deep clone

    // Update main fields with current values from the form
    cleanedObject.priority = formContext.querySelector('[name="priority"]').value;
    cleanedObject.internalCaseNumber = formContext.querySelector('[name="internalCaseNumber"]').value;
    cleanedObject.problemSummary = formContext.querySelector('[name="problemSummary"]').value;
    cleanedObject.problemDescription = formContext.querySelector('[name="problemDescription"]').value;

    // For customer fields, update the current values from the form inputs
    cleanedObject.customFields.forEach(data => {
        console.log(`Processing field ID: ${data.fieldId}`);
        const fieldElement = formContext.querySelector(`[name*="field_${data.fieldId}"]`);

        if (fieldElement) {
            if (fieldElement.tagName === "SELECT" && fieldElement.getAttribute("data-tierselect") === "true") {
                // Get the value from the most child optionset
                const tierSelectElements = formContext.querySelectorAll(`[name^="field_${data.fieldId}"]`);
                data.value = Array.from(tierSelectElements).map(el => el.value).join(' : ');

            } else if (fieldElement.tagName === "SELECT") {
                data.value = fieldElement.value;  // For select dropdowns
            } else {
                data.value = fieldElement.value;  // For text, number, email, etc.
            }
        } else {
            console.error(`Field element not found for field ID: ${data.fieldId}`);
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

    // Case Information Section
    const caseInfoSection = document.createElement("div");
    caseInfoSection.className = "form-section";
    caseInfoSection.innerHTML = "<h3><strong>Case Information:</strong></h3>";
    caseInfoSection.appendChild(createReadOnlyTextField("Company", formJsonData.submitCompanyName));
    caseInfoSection.appendChild(createReadOnlyTextField("Type", formContext.getAttribute("ap_direction").getText()));
    caseInfoSection.appendChild(createReadOnlyTextField("Priority", formJsonData.priority));
    caseInfoSection.appendChild(createReadOnlyTextField("Case#", formJsonData.submitterCaseNumber));
    caseInfoSection.appendChild(createReadOnlyTextField("Date", formJsonData.createdAt));
    caseInfoSection.appendChild(createReadOnlyTextField("Submitted by", `${formJsonData.submittedBy.firstName} ${formJsonData.submittedBy.lastName}`));
    caseInfoSection.appendChild(createReadOnlyTextField("Summary", formJsonData.summary));
    caseInfoSection.appendChild(createReadOnlyTextArea("Description", formJsonData.description));

    form.appendChild(caseInfoSection);

    // Response Section
    const responseSectionMain = document.createElement("div");
    responseSectionMain.className = "form-section";
    responseSectionMain.innerHTML = "<h3><strong>Response:</strong></h3>";
    responseSectionMain.appendChild(createReadOnlyTextField("Company", formJsonData.receiveCompanyName));
    responseSectionMain.appendChild(createReadOnlyTextField("Case #", formJsonData.receiverCaseNumber));
    form.appendChild(responseSectionMain);

    // Existing response section, do not modify
    loadResponses(formContext, webResourceContent).then(() => {
        // Move the responses section to be always below the existing fields in the response section
        const responsesSection = webResourceContent.querySelector(".response-feed");
        if (responsesSection) {
            responseSectionMain.appendChild(responsesSection);
        }
    });

    // Escalation Instructions Section
    const escalationSection = document.createElement("div");
    escalationSection.className = "form-section";
    escalationSection.innerHTML = "<h3><strong>Escalation Instructions:</strong></h3>";
    escalationSection.appendChild(createReadOnlyHtmlField("", formJsonData.escalationInstructions));
    form.appendChild(escalationSection);

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

    loadCollaborationFeed(formContext, webResourceContent);
    setupAddNoteButton(formContext, webResourceContent);

    // Show the collaboration feed section
    const collaborationFeed = webResourceContent.getElementById("collaborationFeed");
    collaborationFeed.style.display = "block";
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
    textArea.rows = 5;

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

    if (label) {
        const labelElement = document.createElement("label");
        labelElement.className = "form-label";
        labelElement.textContent = label;
        inputGroup.appendChild(labelElement);
    }

    const div = document.createElement("div");
    div.className = "readonly-html";
    div.innerHTML = value || "";  // Render HTML content

    inputGroup.appendChild(div);
    return inputGroup;
}

async function loadCollaborationFeed(formContext, webResourceContent) {
    // Get the current record's case ID
    const caseId = formContext.data.entity.getId();
    if (!caseId) return;

    // Fetch data from the ap_tsanetnote table
    const fetchXml = `
        <fetch>
            <entity name="ap_tsanetnote">
                <attribute name="createdon" />
                <attribute name="ap_creatoremail" />
                <attribute name="ap_creatorname" />
                <attribute name="ap_name" />
                <attribute name="ap_description" />
                <attribute name="ap_tsanotecode" />
                <order attribute="ap_tsanotecode" descending="true" />
                <filter>
                    <condition attribute="ap_tsanetcaseid" operator="eq" value="${caseId.replace(/[{}]/g, '')}" />
                </filter>
            </entity>
        </fetch>
    `;

    const result = await Xrm.WebApi.retrieveMultipleRecords("ap_tsanetnote", "?fetchXml=" + encodeURIComponent(fetchXml));
    if (!result.entities || result.entities.length === 0) return;

    // Populate the Collaboration Feed Section
    const collaborationFeedNotes = webResourceContent.getElementById("collaborationFeedNotes");
    collaborationFeedNotes.innerHTML = ""; // Clear existing content

    result.entities.forEach(note => {
        const noteRow = document.createElement("div");
        noteRow.className = "note-row";

        // Content Section
        const noteContent = document.createElement("div");
        noteContent.className = "note-content";
        noteContent.innerHTML = `
           <div class="note-date">Created on: ${new Date(note.createdon).toLocaleString()}</div>
           <div class="note-meta-inline">
               <span class="ms-Icon ms-Icon--EditNote" aria-hidden="true"></span>
                <span>Note created by: ${note.ap_creatorname || "Unknown"}</span>
                <span>${note.ap_creatoremail || "No Email"}</span>
            </div>
            <div class="note-title">${note.ap_name || "No Title"}</div>
            <div class="note-description">${note.ap_description || "-"}</div>
        `;

        // Append content to the row
        noteRow.appendChild(noteContent);

        // Add the row to the feed
        collaborationFeedNotes.appendChild(noteRow);
    });
}

//Add note button click handler
function setupAddNoteButton(formContext, webResourceContent) {
    const addNoteButton = webResourceContent.getElementById("addNoteButton");

    // Check if the event handler is already attached
    if (addNoteButton.addNoteButtonClickHandler) {
        // Remove existing event listener
        addNoteButton.removeEventListener("click", addNoteButton.addNoteButtonClickHandler);
    }

    // Define the handler function within this scope
    addNoteButton.addNoteButtonClickHandler = function addNoteButtonClickHandler() {
        const currentRecordId = formContext.data.entity.getId(); // Get the current record ID
        openQuickCreateForm(currentRecordId, formContext, webResourceContent);
    };

    // Add the event listener
    addNoteButton.addEventListener("click", addNoteButton.addNoteButtonClickHandler);
}

function openQuickCreateForm(recordId, formContext, webResourceContent) {
    const entityFormOptions = {
        entityName: "ap_tsanetnote",
        useQuickCreateForm: true
    };

    const formParameters = {
        "ap_tsanetcaseid": recordId
    };

    Xrm.Navigation.openForm(entityFormOptions, formParameters).then(
        function (success) {
            console.log("Quick create form opened successfully.");
            // Check if the form was saved
            if (success.savedEntityReference) {
                loadCollaborationFeed(formContext, webResourceContent);
            }
        },
        function (error) {
            console.error("Error opening quick create form: ", error);
        }
    );
}

// Function to load and display responses
async function loadResponses(formContext, webResourceContent) {
    // Get the current record's case ID
    const caseId = formContext.data.entity.getId();
    if (!caseId) return;

    // Fetch data from the ap_tsanetresponse table
    const fetchXml = `
        <fetch>
            <entity name="ap_tsanetresponse">
                <attribute name="ap_type" />
                <attribute name="ap_engineername" />
                <attribute name="ap_engineerphone" />
                <attribute name="ap_engineeremail" />
                <attribute name="ap_description" />
                <attribute name="ap_tsaresponsecode" />
                <order attribute="ap_tsaresponsecode" descending="true" />
                <filter>
                    <condition attribute="ap_tsanetcaseid" operator="eq" value="${caseId.replace(/[{}]/g, '')}" />
                </filter>
            </entity>
        </fetch>
    `;

    const result = await Xrm.WebApi.retrieveMultipleRecords("ap_tsanetresponse", "?fetchXml=" + encodeURIComponent(fetchXml));
    if (!result.entities || result.entities.length === 0) return;

    // Create a new section for responses
    const responsesSection = document.createElement("div");
    responsesSection.className = "response-feed";
    responsesSection.innerHTML = "<h3>Responses</h3>";

    result.entities.forEach(response => {
        const responseContainer = document.createElement("div");
        responseContainer.className = "response-container";

        // Add header with type and timestamp
        const header = document.createElement("div");
        header.className = "response-header";
        header.innerHTML = `
            <strong>${response["ap_type@OData.Community.Display.V1.FormattedValue"] || "No Type"}</strong>
            <span class="timestamp">${new Date().toLocaleString()}</span>
        `;
        responseContainer.appendChild(header);

        // Add content section with engineer details and description
        const content = document.createElement("div");
        content.className = "response-content";
        content.innerHTML = `
            <p><strong>Engineer Name:</strong> ${response.ap_engineername || "N/A"}</p>
            <p><strong>Engineer Email:</strong> ${response.ap_engineeremail || "N/A"}</p>
            <p><strong>Engineer Phone:</strong> ${response.ap_engineerphone || "N/A"}</p>
            <p><strong>Description:</strong> ${response.ap_description || "N/A"}</p>
        `;
        responseContainer.appendChild(content);

        responsesSection.appendChild(responseContainer);
    });

    // Append the responses section below the dynamic form container
    webResourceContent.getElementById("dynamicFormContainer").appendChild(responsesSection);
}
