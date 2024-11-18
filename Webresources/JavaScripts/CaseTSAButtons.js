function CloseButton(formContext) {
    // Define the new statecode and statuscode values
    var statecode = 1;
    var statuscode = 2;

    // Get the ID of the current record
    var recordId = formContext.data.entity.getId();

    // Create an object with the updated values
    var updatedFields = {
        "statecode": statecode,
        "statuscode": statuscode
    };

    // Make an asynchronous request to update the record
    Xrm.WebApi.updateRecord(formContext.data.entity.getEntityName(), recordId, updatedFields).then(
        function success(result) {
            console.log("Record updated successfully");
            // Refresh the form
            formContext.data.refresh(true);
        },
        function error(error) {
            console.log("Error updating record: " + error.message);
            // Optionally, you can add code here to handle the error
        }
    );
}
