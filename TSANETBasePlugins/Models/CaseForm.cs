using Newtonsoft.Json;
using System.Collections.Generic;

public class CaseForm
{
    [JsonProperty("documentId")]
    public int DocumentId { get; set; }

    [JsonProperty("internalCaseNumber")]
    public string InternalCaseNumber { get; set; }

    [JsonProperty("receiverInternalCaseNumber")]
    public string RecieverInternalCaseNumber { get; set; }

    [JsonProperty("problemSummary")]
    public string ProblemSummary { get; set; }

    [JsonProperty("problemDescription")]
    public string ProblemDescription { get; set; }

    [JsonProperty("priority")]
    public string priority { get; set; }

    [JsonProperty("adminNote")]
    public string AdminNote { get; set; }

    [JsonProperty("EscalationInstructions")]
    public string EscalationInstructions { get; set; }

    [JsonProperty("testSubmission")]
    public bool TestSubmission { get; set; }

    [JsonProperty("customFields")]
    public List<CustomField> CustomFields { get; set; }

    [JsonProperty("internalNotes")]
    public List<InternalNote> InternalNotes { get; set; }
}
