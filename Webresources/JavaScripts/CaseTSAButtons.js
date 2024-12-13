function CloseButton(formContext) {
    //inactive status, close response
    ProcessAction(formContext, 4, true);
}

function AcceptButton(formContext) {
    //active status, approval response
    ProcessAction(formContext, 1, true);
}

function RequestInfoButton(formContext) {
    //active status, request info response
    ProcessAction(formContext, 2, true);
}

function RejectButton(formContext) {
    //inactive status, reject response
    ProcessAction(formContext, 0, false);
}

function ResponseInfoButton(formContext) {
    //active status, response info response
    ProcessAction(formContext, 3, true);
}

//Common logic to update the case record and create a new tsanetresponse record
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
