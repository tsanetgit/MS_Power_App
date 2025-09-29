function CloseButton(formContext) {
    "use strict";
    //inactive status, close response
    QuickCreateResponse(formContext, 4);
}

function AcceptButton(formContext) {
    "use strict";
    //active status, approval response
    QuickCreateResponse(formContext, 1);
}

function RequestInfoButton(formContext) {
    "use strict";
    //active status, request info response
    QuickCreateResponse(formContext, 2);
}

function RejectButton(formContext) {
    "use strict";
    //inactive status, reject response
    QuickCreateResponse(formContext, 0);
}

function ResponseInfoButton(formContext) {
    "use strict";
    //active status, response info response
    QuickCreateResponse(formContext, 3);
}

function AcceptUpdateButton(formContext) {
    "use strict";
    //active status, response info response
    QuickCreateResponse(formContext, 5);
}

// Function to open quick create form of the response table with prefilled fields
function QuickCreateResponse(formContext, type) {
    "use strict";
    var ownerId = formContext.data.entity.attributes.get("ownerid").getValue()[0].id;
    var recordId = formContext.data.entity.getId().replace(/[{}]/g, "");
    formContext.data.entity.save();
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
                        getCase(formContext);
                    }
                },
                function error(error) {
                    showError(formContext, "Error opening quick create form: " + error.message);
                }
            );
        },
        function error(error) {
            showError(formContext, "Error retrieving owner details: " + error.message);
        }
    );
}

function AddNoteButton(selectedRows) {
    "use strict";
    openQuickCreateForSelectedRow(selectedRows);
}

function openQuickCreateForSelectedRow(selectedRows) {
    "use strict";
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

            }
        },
        function error(error) {

        }
    );
}
