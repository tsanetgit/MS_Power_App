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

            // Create the ap_tsanetresponse record
            var tsanetResponse = {
                "ap_tsanetcaseid@odata.bind": "/ap_tsanetcases(" + recordId + ")",
                "ap_type": 4
            };

            Xrm.WebApi.createRecord("ap_tsanetresponse", tsanetResponse).then(
                function success(result) {
                    console.log("ap_tsanetresponse record created successfully");
                    // Refresh the form
                    formContext.data.refresh(true);
                },
                function error(error) {
                    console.log("Error creating ap_tsanetresponse record: " + error.message);
                    // Optionally, you can add code here to handle the error
                }
            );
        },
        function error(error) {
            console.log("Error updating record: " + error.message);
            // Optionally, you can add code here to handle the error
        }
    );
}
