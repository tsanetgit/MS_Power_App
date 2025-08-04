function onLoad(executionContext) {
    var formContext = executionContext.getFormContext();
    
    // Register event handlers
    formContext.getAttribute("ap_mappingtype").addOnChange(function() {
        handleMappingTypeChange(executionContext);
        handleLookupFieldsVisibility(executionContext);
    });
    
    formContext.getAttribute("ap_attributetype").addOnChange(function() {
        handleLookupFieldsVisibility(executionContext);
    });
    
    // Initialize form state
    handleMappingTypeChange(executionContext);
    handleLookupFieldsVisibility(executionContext);
}

/**
 * Handle visibility of ap_isrequired field based on mapping type
 * @param {Xrm.Events.EventContext} executionContext - The execution context
 */
function handleMappingTypeChange(executionContext) {
    var formContext = executionContext.getFormContext();
    var mappingType = formContext.getAttribute("ap_mappingtype").getValue();

    var showField = mappingType === 120950001;
    
    formContext.getControl("ap_isrequired").setVisible(showField);
}

/**
 * Handle visibility of lookup fields based on attribute type and mapping type
 * @param {Xrm.Events.EventContext} executionContext - The execution context
 */
function handleLookupFieldsVisibility(executionContext) {
    var formContext = executionContext.getFormContext();
    var attributeType = formContext.getAttribute("ap_attributetype").getValue();
    var mappingType = formContext.getAttribute("ap_mappingtype").getValue();
    
    // Show lookup fields when attribute type is 2 and mapping type is 120950000
    var showLookupFields = attributeType === 2 && mappingType === 120950000;
    
    formContext.getControl("ap_lookupentityname").setVisible(showLookupFields);
    formContext.getControl("ap_lookupsearchattribute").setVisible(showLookupFields);
}