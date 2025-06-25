var TSA = TSA || {};

(function () {
    "use strict";

    // Constants
    TSA.UPLOAD_NOTIFICATION_ID = "fileUploadStatus";
    TSA.UPLOAD_IN_PROGRESS_MESSAGE = "Upload in progress";
    TSA.UPLOAD_FAILED_MESSAGE = "Upload failed, please contact your system administrator.";
    TSA.CHECK_INTERVAL_MS = 5000;

    // Variable to store the interval ID
    let annotationCheckIntervalId = null;

    // Function to wait for web resource element
    TSA.waitForWebResourceElement = function (formContext, webResourceName, elementId, callback, checkFrequency = 500, timeout = 5000) {
        const startTime = Date.now();
        const interval = setInterval(function () {
            try {
                const webResourceControl = parent.Xrm.Page.getControl(webResourceName);
                if (webResourceControl) {
                    const webResourceContent = webResourceControl.getObject().contentDocument;
                    const element = webResourceContent.getElementById(elementId);  // Directly access the element
                    if (element) {
                        clearInterval(interval);
                        callback();
                    }
                }
            } catch (e) {
                TSA.showError(formContext, "An error occurred: " + e.message);
            }

            if (Date.now() - startTime > timeout) {
                clearInterval(interval);
                TSA.showError(formContext, "Element not found within the timeout period.");
            }
        }, checkFrequency);
    };

    // Helper function to disable submit
    TSA.disableButton = function (disable, webResource) {
        // Disable the submit button in the dynamic form
        const webResourceControl = parent.Xrm.Page.getControl(webResource);
        const webResourceContent = webResourceControl.getObject().contentDocument;
        const submitButton = webResourceContent.querySelector("button[type='submit']");
        if (submitButton) {
            submitButton.disabled = disable;
        }
    };

    // Helper that shows an error and removes it after 5 seconds
    TSA.showError = function (formContext, message) {
        formContext.ui.setFormNotification(message, "ERROR", "tsaerror");

        // Set a timeout to remove the error notification after 5 seconds (5000 milliseconds)
        setTimeout(function () {
            formContext.ui.clearFormNotification("tsaerror");
        }, 5000);
    };

    // Helper function to save the object to a field
    TSA.saveToFormField = function (fieldName, formData, formContext) {
        const newValue = JSON.stringify(formData);
        const currentValue = formContext.getAttribute(fieldName).getValue();

        // Only update the field if the new value is different from the current value
        if (currentValue !== newValue) {
            formContext.getAttribute(fieldName).setValue(newValue);
            formContext.getAttribute(fieldName).setSubmitMode("dirty");
            formContext.getAttribute(fieldName).fireOnChange();
        }
    };

    // Function to show warning message
    TSA.showWarningMessage = function (formContext, message) {
        formContext.ui.setFormNotification(message, "WARNING", "statuscodeWarning");
    };

    // Function to clear warning message
    TSA.clearWarningMessage = function (formContext) {
        formContext.ui.clearFormNotification("statuscodeWarning");
    };

    // Helper function to map priority from string to option set value
    TSA.getPriorityValue = function (priority) {
        switch (priority.toLowerCase()) {
            case "low": return 3;
            case "medium": return 2;
            case "high": return 1;
            default: return null;
        }
    };

    // Helper function to get response type value
    TSA.getResponseTypeValue = function (type) {
        switch (type.toLowerCase()) {
            case 'approval':
                return 1;
            case 'rejection':
                return 0;
            case 'information_request':
                return 2;
            case 'information_response':
                return 3;
            default:
                return null;
        }
    };

    // Function to refresh read-only form
    TSA.refreshReadOnlyForm = function (formContext) {
        formContext.data.refresh(true);
        const formJsonField = formContext.getAttribute("ap_formjson").getValue();

        if (formJsonField) {
            // If `ap_formjson` contains data, parse it and build the read-only form
            const formJsonData = JSON.parse(formJsonField);
            buildReadOnlyForm(formJsonData, formContext);
        } else {
            TSA.showError(formContext, "No form JSON data available to refresh the read-only form.");
        }
    };


    TSA.statusWarningLogic = function (formContext) {
        const statusCode = formContext.getAttribute("statuscode").getValue();
        const direction = formContext.getAttribute("ap_direction").getValue();

        // Show warning if statuscode is 1
        if (statusCode === 1 && direction === 0) {
            TSA.showWarningMessage(formContext, "Response needed (Accept, request information, or Reject)​");
        }
        else {
            TSA.clearWarningMessage(formContext);
        }
    };

    // Helper that shows a Success notification and removes it after 10 seconds
    TSA.showSuccess = function (formContext, message) {
        formContext.ui.setFormNotification(message, "INFO", "tsasuccess");

        // Set a timeout to remove the notification after 10 seconds (10000 milliseconds)
        setTimeout(function () {
            formContext.ui.clearFormNotification("tsasuccess");
        }, 10000);
    };

    // Function to show "Upload in progress" notification
    TSA.showUploadInProgressNotification = function (formContext) {
        formContext.ui.setFormNotification(
            TSA.UPLOAD_IN_PROGRESS_MESSAGE,
            "INFO",
            TSA.UPLOAD_NOTIFICATION_ID
        );
    };

    // Function to show "Upload failed" notification
    TSA.showUploadFailedNotification = function (formContext) {
        formContext.ui.setFormNotification(
            TSA.UPLOAD_FAILED_MESSAGE,
            "ERROR",
            TSA.UPLOAD_NOTIFICATION_ID
        );
    };

    // Function to clear upload notification
    TSA.clearUploadNotification = function (formContext) {
        formContext.ui.clearFormNotification(TSA.UPLOAD_NOTIFICATION_ID);
    };

    // Function to check annotation status
    TSA.checkAnnotationStatus = function (formContext) {
        const caseId = formContext.data.entity.getId().replace(/[{}]/g, "");

        // Query to fetch annotations with subject = '%TOPROCESS%' for this case
        const query = `?$filter=_objectid_value eq ${caseId} and subject eq '%TOPROCESS%'`;

        return Xrm.WebApi.retrieveMultipleRecords("annotation", query).then(
            function success(result) {
                if (result.entities.length === 0) {
                    // No more annotations with %TOPROCESS% subject - processing complete
                    TSA.clearUploadNotification(formContext);
                    return { status: "completed" };
                } else {
                    // Check if any of the notes have "Upload Failed!" in the text
                    const failedUploads = result.entities.filter(
                        note => note.notetext && note.notetext.includes("Upload failed!")
                    );

                    if (failedUploads.length > 0) {
                        // At least one upload failed
                        TSA.showUploadFailedNotification(formContext);
                        return { status: "failed" };
                    }

                    // Uploads still in progress
                    return { status: "inProgress" };
                }
            },
            function error(error) {
                TSA.showError(formContext, "Error checking annotation status:", error.message);
                return { status: "error", message: error.message };
            }
        );
    };

    // Function to start periodic check for annotation status
    TSA.startAnnotationStatusCheck = function (formContext) {
        // Clear any existing interval first
        TSA.stopAnnotationStatusCheck();

        // Start a new interval
        annotationCheckIntervalId = setInterval(function () {
            TSA.checkAnnotationStatus(formContext).then(result => {
                if (result.status === "completed" || result.status === "failed") {
                    // Stop checking if processing is complete or failed
                    TSA.stopAnnotationStatusCheck();
                    getCase(formContext);
                }
            });
        }, TSA.CHECK_INTERVAL_MS);

        // Also store the interval ID in a form-level variable to ensure it can be cleared
        // when the form is closed or refreshed
        if (formContext && formContext.context && formContext.context.data) {
            formContext.context.data._annotationCheckIntervalId = annotationCheckIntervalId;
        }
    };

    // Function to stop periodic check
    TSA.stopAnnotationStatusCheck = function () {
        if (annotationCheckIntervalId) {
            clearInterval(annotationCheckIntervalId);
            annotationCheckIntervalId = null;
        }
    };

    // Function to initialize upload notification monitoring on form load
    TSA.initializeUploadNotificationMonitoring = function (formContext) {
        // Check if there are any pending uploads
        TSA.checkAnnotationStatus(formContext).then(result => {
            if (result.status === "inProgress") {
                TSA.showUploadInProgressNotification(formContext);
                TSA.startAnnotationStatusCheck(formContext);
            } else if (result.status === "failed") {
                TSA.showUploadFailedNotification(formContext);
            }
        });
    };

    TSA.updateTsaCaseStatus = function (formContext, statusString) {
        let stateCode = 0;  // Default
        let statusCode = 1;  

        // Convert the status string to lowercase for case-insensitive comparison
        const status = statusString.toLowerCase();

        switch (status) {
            case "rejected":
                stateCode = 1;  // Inactive
                statusCode = 120950002;  // Rejected
                break;
            case "closed":
                stateCode = 1;  // Inactive
                statusCode = 2;  // Closed
                break;
            case "information":
                stateCode = 0;  // Active
                statusCode = 120950001;  // Information
                break;
            case "accepted":
                stateCode = 0;  // Active
                statusCode = 120950003;  // Accepted
                break;
            default:
                break;
        }

        // Update the form status
        formContext.getAttribute("statecode").setValue(stateCode);
        formContext.getAttribute("statuscode").setValue(statusCode);
        formContext.data.entity.save();
    };

    // Add to the window object for backward compatibility
    // For backward compatibility, create references to the global namespace
    // This allows existing code to continue working without modification
    window.waitForElementsToExist = function () { return TSA.waitForElementsToExist.apply(this, arguments); };
    window.waitForWebResourceElement = function () { return TSA.waitForWebResourceElement.apply(this, arguments); };
    window.disableButton = function () { return TSA.disableButton.apply(this, arguments); };
    window.showError = function () { return TSA.showError.apply(this, arguments); };
    window.saveToFormField = function () { return TSA.saveToFormField.apply(this, arguments); };
    window.showWarningMessage = function () { return TSA.showWarningMessage.apply(this, arguments); };
    window.clearWarningMessage = function () { return TSA.clearWarningMessage.apply(this, arguments); };
    window.getPriorityValue = function () { return TSA.getPriorityValue.apply(this, arguments); };
    window.getResponseTypeValue = function () { return TSA.getResponseTypeValue.apply(this, arguments); };
    window.refreshReadOnlyForm = function () { return TSA.refreshReadOnlyForm.apply(this, arguments); };
    window.showSuccess = function () { return TSA.showSuccess.apply(this, arguments); };
    window.showUploadInProgressNotification = function () { return TSA.showUploadInProgressNotification.apply(this, arguments); };
    window.showUploadFailedNotification = function () { return TSA.showUploadFailedNotification.apply(this, arguments); };
    window.clearUploadNotification = function () { return TSA.clearUploadNotification.apply(this, arguments); };
    window.checkAnnotationStatus = function () { return TSA.checkAnnotationStatus.apply(this, arguments); };
    window.startAnnotationStatusCheck = function () { return TSA.startAnnotationStatusCheck.apply(this, arguments); };
    window.stopAnnotationStatusCheck = function () { return TSA.stopAnnotationStatusCheck.apply(this, arguments); };
    window.initializeUploadNotificationMonitoring = function () { return TSA.initializeUploadNotificationMonitoring.apply(this, arguments); };
    window.statusWarningLogic = function () { return TSA.statusWarningLogic.apply(this, arguments); };
    window.updateTsaCaseStatus = function () { return TSA.updateTsaCaseStatus.apply(this, arguments); };
})();