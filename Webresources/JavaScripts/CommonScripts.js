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