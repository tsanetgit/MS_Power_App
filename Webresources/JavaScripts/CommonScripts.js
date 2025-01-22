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

// Helper function to disableSubmit
function disableButton(disable, webResource) {
    // Disable the submit button in the dynamic form
    const webResourceControl = parent.Xrm.Page.getControl(webResource);
    const webResourceContent = webResourceControl.getObject().contentDocument;
    const submitButton = webResourceContent.querySelector("button[type='submit']");
    if (submitButton) {
        submitButton.disabled = disable;
    }
}

// Helper that shows an error and removes it after 5 seconds
function showError(formContext, message) {
    formContext.ui.setFormNotification(message, "ERROR", "tsaerror");

    // Set a timeout to remove the error notification after 5 seconds (5000 milliseconds)
    setTimeout(function () {
        formContext.ui.clearFormNotification("tsaerror");
    }, 5000);
}

// Helper function to save the object to a field
function saveToFormField(fieldName, formData, formContext) {
    const newValue = JSON.stringify(formData);
    const currentValue = formContext.getAttribute(fieldName).getValue();

    // Only update the field if the new value is different from the current value
    if (currentValue !== newValue) {
        formContext.getAttribute(fieldName).setValue(newValue);
        formContext.getAttribute(fieldName).setSubmitMode("dirty");
        formContext.getAttribute(fieldName).fireOnChange();
    }
}

// Function to show warning message
function showWarningMessage(formContext, message) {
    formContext.ui.setFormNotification(message, "WARNING", "statuscodeWarning");
}

// Function to clear warning message
function clearWarningMessage(formContext) {
    formContext.ui.clearFormNotification("statuscodeWarning");
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

function refreshReadOnlyForm(formContext) {
    const formJsonField = formContext.getAttribute("ap_formjson").getValue();

    if (formJsonField) {
        // If `ap_formjson` contains data, parse it and build the read-only form
        const formJsonData = JSON.parse(formJsonField);
        buildReadOnlyForm(formJsonData, formContext);
    } else {
        console.error("No form JSON data available to refresh the read-only form.");
    }
}

