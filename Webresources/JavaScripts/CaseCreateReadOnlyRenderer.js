"use strict";

/**
 * CaseCreateReadOnlyRenderer
 * Self-contained class for rendering read-only case forms in a Dynamics 365 webresource.
 * Handles iframe reload/tab switching in Customer Service Workspace.
 */
class CaseCreateReadOnlyRenderer {
    constructor(config = {}) {
        this.config = {
            jsonFieldName: config.jsonFieldName || "ap_formjson",
            containerId: config.containerId || "dynamicFormContainer",
            companySearchId: config.companySearchId || "company-search",
            collaborationFeedId: config.collaborationFeedId || "collaborationFeed",
            collaborationNotesId: config.collaborationNotesId || "collaborationFeedNotes",
            addNoteButtonId: config.addNoteButtonId || "addNoteButton",
            uploadButtonId: config.uploadButtonId || "uploadAttachmentButton"
        };

        this.logPrefix = "[CaseCreateWR]";
        this.initialized = false;
        this.visibilityChangeHandler = null;
        this.currentRecordContext = null;
    }

    /**
     * Initialize the renderer: wire events and trigger initial refresh
     */
    init() {
        console.log(`${this.logPrefix} init() called`);

        // Prevent duplicate initialization
        if (this.initialized) {
            console.log(`${this.logPrefix} Already initialized, skipping`);
            return;
        }

        // Check if we have a record ID - skip if this is a create form
        const params = new URLSearchParams(window.location.search);
        const id = params.get("id") || params.get("data");

        if (!id) {
            console.log(`${this.logPrefix} No record ID found - skipping renderer (create form scenario)`);
            return;
        }

        // Wire up visibility change handler (only once)
        this.visibilityChangeHandler = () => {
            if (document.visibilityState === "visible") {
                console.log(`${this.logPrefix} Document became visible, triggering refresh`);
                this.refresh();
            }
        };
        document.addEventListener("visibilitychange", this.visibilityChangeHandler);

        // Expose global refresh method for parent window
        window.refreshReadOnlyForm = () => {
            console.log(`${this.logPrefix} Manual refresh triggered via window.refreshReadOnlyForm()`);
            this.refresh();
        };

        this.initialized = true;

        // Initial refresh
        this.refresh();
    }

    /**
     * Main refresh method: retrieves record context, fetches JSON, and renders form
     */
    async refresh() {
        console.log(`${this.logPrefix} refresh() called`);

        try {
            // Get record context from query string
            this.currentRecordContext = this.getRecordContextFromQueryString();
            console.log(`${this.logPrefix} Record context:`, this.currentRecordContext);

            if (!this.currentRecordContext.id || !this.currentRecordContext.entityLogicalName) {
                this.showMessage("Missing record context parameters. Ensure the webresource is configured to pass record object type code and unique identifier.");
                return;
            }

            // Retrieve JSON from Dataverse
            const jsonData = await this.retrieveJson(
                this.currentRecordContext.entityLogicalName,
                this.currentRecordContext.id
            );

            if (!jsonData) {
                this.showMessage("No form JSON available for this record.");
                return;
            }

            console.log(`${this.logPrefix} JSON retrieved, length: ${JSON.stringify(jsonData).length}`);

            // Render the read-only form
            await this.renderReadOnlyForm(jsonData);

        } catch (error) {
            console.error(`${this.logPrefix} Error during refresh:`, error);
            this.showMessage(`Error loading form: ${error.message || error}`);
        }
    }

    /**
     * Parse query string to extract id and entity type name
     * @returns {{ id: string, entityLogicalName: string }}
     */
    getRecordContextFromQueryString() {
        const params = new URLSearchParams(window.location.search);
        let id = params.get("id") || params.get("data");
        const etn = params.get("etn") || params.get("typename");

        // Normalize GUID: remove braces if present
        if (id) {
            id = id.replace(/[{}]/g, "");
        }

        return {
            id: id,
            entityLogicalName: etn
        };
    }

    /**
     * Retrieve JSON field from Dataverse
     * @param {string} entityName - Entity logical name
     * @param {string} id - Record GUID (without braces)
     * @returns {Promise<object|null>} Parsed JSON object or null
     */
    async retrieveJson(entityName, id) {
        try {
            if (!parent.Xrm || !parent.Xrm.WebApi) {
                throw new Error("parent.Xrm.WebApi is not available");
            }

            const selectQuery = `?$select=${this.config.jsonFieldName}`;
            const record = await parent.Xrm.WebApi.retrieveRecord(entityName, id, selectQuery);

            const jsonField = record[this.config.jsonFieldName];

            if (!jsonField) {
                console.log(`${this.logPrefix} ${this.config.jsonFieldName} is empty or null`);
                return null;
            }

            // Parse JSON
            const parsedJson = JSON.parse(jsonField);
            return parsedJson;

        } catch (error) {
            if (error.name === "SyntaxError") {
                throw new Error(`Invalid JSON in ${this.config.jsonFieldName}: ${error.message}`);
            }
            throw error;
        }
    }

    /**
     * Render the read-only form into the container
     * @param {object} formJsonData - Parsed JSON data
     */
    async renderReadOnlyForm(formJsonData) {
        console.log(`${this.logPrefix} renderReadOnlyForm() called`);

        const container = document.getElementById(this.config.containerId);
        if (!container) {
            throw new Error(`Container element #${this.config.containerId} not found`);
        }

        container.innerHTML = ""; // Clear existing content

        const form = document.createElement("form");
        form.className = "dynamic-form";

        // Case Information Section
        const caseInfoSection = document.createElement("div");
        caseInfoSection.className = "form-section";
        caseInfoSection.innerHTML = "<h3><strong>Case Information:</strong></h3>";
        caseInfoSection.appendChild(this.createReadOnlyTextField("Company", formJsonData.submitCompanyName));
        
        // Get direction text from formJsonData if available
        const directionText = formJsonData.direction || formJsonData.caseDirection || "N/A";
        caseInfoSection.appendChild(this.createReadOnlyTextField("Type", directionText));
        
        caseInfoSection.appendChild(this.createReadOnlyTextField("Priority", formJsonData.priority));
        caseInfoSection.appendChild(this.createReadOnlyTextField("Case#", formJsonData.submitterCaseNumber));
        caseInfoSection.appendChild(this.createReadOnlyTextField("Date", formJsonData.createdAt));
        
        const submittedByName = formJsonData.submittedBy 
            ? `${formJsonData.submittedBy.firstName || ""} ${formJsonData.submittedBy.lastName || ""}`.trim()
            : "N/A";
        caseInfoSection.appendChild(this.createReadOnlyTextField("Submitted by", submittedByName));
        
        caseInfoSection.appendChild(this.createReadOnlyTextField("Summary", formJsonData.summary));
        caseInfoSection.appendChild(this.createReadOnlyTextArea("Description", formJsonData.description));
        form.appendChild(caseInfoSection);

        // Submitter Section
        if (formJsonData.submitterContactDetails) {
            const submitterSectionMain = document.createElement("div");
            submitterSectionMain.className = "form-section";
            submitterSectionMain.innerHTML = "<h3><strong>Submitter Contact Details:</strong></h3>";
            submitterSectionMain.appendChild(this.createReadOnlyTextField("Name", formJsonData.submitterContactDetails.name));
            submitterSectionMain.appendChild(this.createReadOnlyTextField("E-mail", formJsonData.submitterContactDetails.email));
            submitterSectionMain.appendChild(this.createReadOnlyTextField("Phone", formJsonData.submitterContactDetails.phone));
            form.appendChild(submitterSectionMain);
        }

        // Response Section
        const responseSectionMain = document.createElement("div");
        responseSectionMain.className = "form-section";
        responseSectionMain.innerHTML = "<h3><strong>Response:</strong></h3>";
        responseSectionMain.appendChild(this.createReadOnlyTextField("Company", formJsonData.receiveCompanyName));
        responseSectionMain.appendChild(this.createReadOnlyTextField("Case #", formJsonData.receiverCaseNumber));
        form.appendChild(responseSectionMain);

        // Load responses and append to response section
        try {
            await this.loadResponses(responseSectionMain);
        } catch (error) {
            console.error(`${this.logPrefix} Error loading responses:`, error);
        }

        // Escalation Instructions Section
        if (formJsonData.escalationInstructions) {
            const escalationSection = document.createElement("div");
            escalationSection.className = "form-section";
            escalationSection.innerHTML = "<h3><strong>Escalation Instructions:</strong></h3>";
            escalationSection.appendChild(this.createReadOnlyHtmlField("", formJsonData.escalationInstructions));
            form.appendChild(escalationSection);
        }

        // Custom Fields
        if (formJsonData.customFields && formJsonData.customFields.length > 0) {
            const customFields = this.groupBy(formJsonData.customFields, "section");
            for (const section in customFields) {
                // Map section names to user-friendly display names
                let displaySectionName = section;
                if (section.toUpperCase() === "CONTACT_SECTION") {
                    displaySectionName = "Contact";
                } else if (section.toUpperCase() === "COMMON_CUSTOMER_SECTION") {
                    displaySectionName = "Customer Details";
                } else if (section.toUpperCase() === "PROBLEM_SECTION") {
                    displaySectionName = "Problem Details";
                }

                const sectionGroup = document.createElement("div");
                sectionGroup.className = "form-section";
                sectionGroup.innerHTML = `<h3>${displaySectionName}</h3>`;

                customFields[section].forEach(field => {
                    sectionGroup.appendChild(this.createReadOnlyTextField(field.fieldName || field.label, field.value));
                });
                form.appendChild(sectionGroup);
            }
        }

        container.appendChild(form);

        // Load collaboration feed
        try {
            await this.loadCollaborationFeed();
            this.setupAddNoteButton();
            this.registerUploadButton();
        } catch (error) {
            console.error(`${this.logPrefix} Error loading collaboration features:`, error);
        }

        // Show the collaboration feed section
        const collaborationFeed = document.getElementById(this.config.collaborationFeedId);
        if (collaborationFeed) {
            collaborationFeed.style.display = "block";
        }

        console.log(`${this.logPrefix} Read-only form rendered successfully`);
    }

    /**
     * Display a message in the container (for errors or info)
     * @param {string} message
     */
    showMessage(message) {
        const container = document.getElementById(this.config.containerId);
        if (container) {
            container.innerHTML = `<div class="info-message">${message}</div>`;
        }
        console.log(`${this.logPrefix} ${message}`);
    }

    // ========== HELPER FUNCTIONS (migrated from CaseTSA.js) ==========

    /**
     * Create read-only text field
     */
    createReadOnlyTextField(label, value) {
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

    /**
     * Create read-only text area
     */
    createReadOnlyTextArea(label, value) {
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

    /**
     * Create read-only HTML field
     */
    createReadOnlyHtmlField(label, value) {
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
        div.innerHTML = value || "";

        inputGroup.appendChild(div);
        return inputGroup;
    }

    /**
     * Group array by key
     */
    groupBy(arr, key) {
        return arr.reduce((group, item) => {
            const section = item[key];
            group[section] = group[section] || [];
            group[section].push(item);
            return group;
        }, {});
    }

    /**
     * Load responses from ap_tsanetresponse table
     */
    async loadResponses(responseSectionMain) {
        if (!this.currentRecordContext || !this.currentRecordContext.id) {
            console.log(`${this.logPrefix} No record context for loading responses`);
            return;
        }

        const caseId = this.currentRecordContext.id;

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
                        <condition attribute="ap_tsanetcaseid" operator="eq" value="${caseId}" />
                    </filter>
                </entity>
            </fetch>
        `;

        try {
            const result = await parent.Xrm.WebApi.retrieveMultipleRecords("ap_tsanetresponse", `?fetchXml=${encodeURIComponent(fetchXml)}`);
            
            if (!result.entities || result.entities.length === 0) {
                console.log(`${this.logPrefix} No responses found`);
                return;
            }

            // Create responses section
            const responsesSection = document.createElement("div");
            responsesSection.className = "response-feed";
            responsesSection.innerHTML = "<h3>Responses</h3>";

            result.entities.forEach(response => {
                const responseContainer = document.createElement("div");
                responseContainer.className = "response-container";

                const header = document.createElement("div");
                header.className = "response-header";
                header.innerHTML = `
                    <strong>${response["ap_type@OData.Community.Display.V1.FormattedValue"] || "No Type"}</strong>
                    <span class="timestamp">${new Date().toLocaleString()}</span>
                `;
                responseContainer.appendChild(header);

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

            responseSectionMain.appendChild(responsesSection);
            console.log(`${this.logPrefix} Loaded ${result.entities.length} response(s)`);

        } catch (error) {
            console.error(`${this.logPrefix} Error loading responses:`, error);
        }
    }

    /**
     * Load collaboration feed (notes)
     */
    async loadCollaborationFeed() {
        if (!this.currentRecordContext || !this.currentRecordContext.id) {
            console.log(`${this.logPrefix} No record context for loading collaboration feed`);
            return;
        }

        const caseId = this.currentRecordContext.id;

        const fetchXml = `
            <fetch>
                <entity name="ap_tsanetnote">
                    <attribute name="createdon" />
                    <attribute name="ap_creatoremail" />
                    <attribute name="ap_creatorname" />
                    <attribute name="ap_name" />
                    <attribute name="ap_description" />
                    <attribute name="ap_tsanotecode" />
                    <order attribute="createdon" descending="true" />
                    <filter>
                        <condition attribute="ap_tsanetcaseid" operator="eq" value="${caseId}" />
                    </filter>
                </entity>
            </fetch>
        `;

        try {
            const result = await parent.Xrm.WebApi.retrieveMultipleRecords("ap_tsanetnote", `?fetchXml=${encodeURIComponent(fetchXml)}`);
            
            const collaborationFeedNotes = document.getElementById(this.config.collaborationNotesId);
            if (!collaborationFeedNotes) {
                console.log(`${this.logPrefix} Collaboration notes container not found`);
                return;
            }

            collaborationFeedNotes.innerHTML = "";

            if (!result.entities || result.entities.length === 0) {
                console.log(`${this.logPrefix} No collaboration notes found`);
                return;
            }

            result.entities.forEach(note => {
                const noteRow = document.createElement("div");
                noteRow.className = "note-row";

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

                noteRow.appendChild(noteContent);
                collaborationFeedNotes.appendChild(noteRow);
            });

            console.log(`${this.logPrefix} Loaded ${result.entities.length} collaboration note(s)`);

        } catch (error) {
            console.error(`${this.logPrefix} Error loading collaboration feed:`, error);
        }
    }

    /**
     * Setup add note button
     */
    setupAddNoteButton() {
        const addNoteButton = document.getElementById(this.config.addNoteButtonId);
        if (!addNoteButton) {
            console.log(`${this.logPrefix} Add note button not found`);
            return;
        }

        // Remove existing handler if present (prevent duplicates)
        if (addNoteButton.addNoteButtonClickHandler) {
            addNoteButton.removeEventListener("click", addNoteButton.addNoteButtonClickHandler);
        }

        addNoteButton.addNoteButtonClickHandler = () => {
            this.openQuickCreateForm();
        };

        addNoteButton.addEventListener("click", addNoteButton.addNoteButtonClickHandler);
        console.log(`${this.logPrefix} Add note button configured`);
    }

    /**
     * Open quick create form for notes
     */
    openQuickCreateForm() {
        if (!this.currentRecordContext || !this.currentRecordContext.id) {
            console.error(`${this.logPrefix} No record context for creating note`);
            return;
        }

        const recordId = `{${this.currentRecordContext.id}}`;

        const entityFormOptions = {
            entityName: "ap_tsanetnote",
            useQuickCreateForm: true
        };

        const formParameters = {
            "ap_tsanetcaseid": recordId
        };

        parent.Xrm.Navigation.openForm(entityFormOptions, formParameters).then(
            (success) => {
                if (success.savedEntityReference) {
                    console.log(`${this.logPrefix} Note created, reloading collaboration feed`);
                    this.loadCollaborationFeed();
                }
            },
            (error) => {
                console.error(`${this.logPrefix} Error opening quick create form:`, error);
            }
        );
    }

    /**
     * Register upload button (placeholder - depends on external functions)
     */
    registerUploadButton() {
        const uploadButton = document.getElementById(this.config.uploadButtonId);
        if (!uploadButton) {
            console.log(`${this.logPrefix} Upload button not found`);
            return;
        }

        // Initially hide until we check configuration
        uploadButton.style.display = "none";

        // Note: The original implementation depends on getAttachmentConfig and createNoteWithFile
        // which likely require formContext. For now, we'll leave this as a placeholder.
        // Future work: migrate attachment logic to be webresource-compatible.

        console.log(`${this.logPrefix} Upload button found but not fully configured (needs migration)`);
    }

    /**
     * Cleanup (if needed)
     */
    destroy() {
        if (this.visibilityChangeHandler) {
            document.removeEventListener("visibilitychange", this.visibilityChangeHandler);
        }
        this.initialized = false;
        console.log(`${this.logPrefix} Renderer destroyed`);
    }
}

// Expose globally for instantiation from HTML
window.CaseCreateReadOnlyRenderer = CaseCreateReadOnlyRenderer;