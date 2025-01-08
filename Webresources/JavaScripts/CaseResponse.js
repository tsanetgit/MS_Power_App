function onLoad(executionContext) {
    var formContext = executionContext.getFormContext();
    var apType = formContext.getAttribute("ap_type").getValue();

    // Set visibility for ap_description
    var apDescriptionVisible = [0, 1, 2, 3, 5].includes(apType);
    formContext.getControl("ap_description").setVisible(apDescriptionVisible);

    // Set visibility for ap_internalcasenumber
    var apInternalCaseNumberVisible = [1, 5].includes(apType);
    formContext.getControl("ap_internalcasenumber").setVisible(apInternalCaseNumberVisible);
}
