function CloseButton(formContext) {
    //inactive status, close response
    QuickCreateResponse(formContext, 4);
}

function AcceptButton(formContext) {
    //active status, approval response
    QuickCreateResponse(formContext, 1);
}

function RequestInfoButton(formContext) {
    //active status, request info response
    QuickCreateResponse(formContext, 2);
}

function RejectButton(formContext) {
    //inactive status, reject response
    QuickCreateResponse(formContext, 0);
}

function ResponseInfoButton(formContext) {
    //active status, response info response
    QuickCreateResponse(formContext, 3);
}

function AcceptUpdateButton(formContext) {
    //active status, response info response
    QuickCreateResponse(formContext, 5);
}

//Common logic to update the case record and create a new tsanetresponse record - NOT USED
function ProcessAction(formContext, type, requireDescription) {
    var description = "";
    if (requireDescription) {
        // Show a prompt dialog to get the description from the user
        userIpnut = prompt("Please enter a description:");

        // If the user clicks cancel, userInput will be null
        if (userIpnut === null) {
            console.log("Action cancelled by the user.");
            return;
        }
        else {
            description = userIpnut;
        }
    } else {
        // Show a confirm dialog to get the user's confirmation
        var userConfirmed = confirm("Do you want to proceed?");

        // If the user clicks cancel, userConfirmed will be false
        if (!userConfirmed) {
            console.log("Action cancelled by the user.");
            return;
        }
    }

    // Get the ID of the current record
    var recordId = formContext.data.entity.getId().replace(/[{}]/g, "");

    // Create the ap_tsanetresponse record
    var tsanetResponse = {
        "ap_tsanetcaseid@odata.bind": "/ap_tsanetcases(" + recordId + ")",
        "ap_description": description,
        "ap_type": type
    };

    Xrm.WebApi.createRecord("ap_tsanetresponse", tsanetResponse).then(
        function success(result) {
            console.log("ap_tsanetresponse record created successfully");
            formContext.data.refresh(true);

            //// Make an asynchronous request to update the record
            //Xrm.WebApi.updateRecord(formContext.data.entity.getEntityName(), recordId, updatedFields).then(
            //    function success(result) {
            //        console.log("Record updated successfully");
            //        // Refresh the form
                    
            //    },
            //    function error(error) {
            //        console.log("Error updating record: " + error.message);
            //    }
            //);

        },
        function error(error) {
            console.log("Error creating ap_tsanetresponse record: " + error.message);
            alert(error.message);
        }
    );
}
// Function to open quick create form of the response table with prefilled fields
function QuickCreateResponse(formContext, type) {
    var ownerId = formContext.data.entity.attributes.get("ownerid").getValue()[0].id;
    var recordId = formContext.data.entity.getId().replace(/[{}]/g, "");

    Xrm.WebApi.retrieveRecord("systemuser", ownerId, "?$select=fullname,address1_telephone1,internalemailaddress").then(
        function success(result) {
            var quickCreateData = {
                "ap_type": type,
                "ap_engineername": result.fullname,
                "ap_engineerphone": result.address1_telephone1,
                "ap_engineeremail": result.internalemailaddress,
                "ap_tsanetcaseid": recordId
            };

            var entityFormOptions = {
                entityName: "ap_tsanetresponse",
                useQuickCreateForm: true
            };

            Xrm.Navigation.openForm(entityFormOptions, quickCreateData).then(
                function success(result) {
                    if (result.savedEntityReference) {
                        console.log("Quick create form saved successfully");
                        refreshReadOnlyForm(formContext);
                    }
                },
                function error(error) {
                    console.log("Error opening quick create form: " + error.message);
                    alert(error.message);
                }
            );
        },
        function error(error) {
            console.log("Error retrieving owner details: " + error.message);
            alert(error.message);
        }
    );
}

function AddNoteButton(selectedRows) {
    openQuickCreateForSelectedRow(selectedRows);
}

function openQuickCreateForSelectedRow(selectedRows) {
    // Check if exactly one row is selected
    if (selectedRows.length !== 1) {
        alert("Please select exactly one row.");
        return;
    }

    // Get the ID of the selected row
    var selectedRowId = selectedRows[0].replace(/[{}]/g, "");

    // Define the data to prefill in the quick create form
    var quickCreateData = {
        "ap_tsanetcaseid": selectedRowId
    };

    // Define the entity form options
    var entityFormOptions = {
        entityName: "ap_tsanetnote",
        useQuickCreateForm: true
    };

    // Open the quick create form
    Xrm.Navigation.openForm(entityFormOptions, quickCreateData).then(
        function success(result) {
            if (result.savedEntityReference) {
                console.log("Quick create form saved successfully");
            }
        },
        function error(error) {
            console.log("Error opening quick create form: " + error.message);
            alert(error.message);
        }
    );
}
