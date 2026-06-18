using Newtonsoft.Json;
using System;
using System.Collections.Generic;

public class Case
{
    [JsonProperty("id")]
    public int Id { get; set; }

    [JsonProperty("submitCompanyName")]
    public string SubmitCompanyName { get; set; }

    [JsonProperty("submitCompanyId")]
    public int SubmitCompanyId { get; set; }

    [JsonProperty("submitterCaseNumber")]
    public string SubmitterCaseNumber { get; set; }

    [JsonProperty("receiveCompanyName")]
    public string ReceiveCompanyName { get; set; }

    [JsonProperty("receiveCompanyId")]
    public int ReceiveCompanyId { get; set; }

    [JsonProperty("receiverCaseNumber")]
    public string ReceiverCaseNumber { get; set; }

    [JsonProperty("summary")]
    public string Summary { get; set; }

    [JsonProperty("description")]
    public string Description { get; set; }

    [JsonProperty("priority")]
    public string Priority { get; set; }

    [JsonProperty("status")]
    public string Status { get; set; }

    [JsonProperty("token")]
    public string Token { get; set; }

    [JsonProperty("createdAt")]
    public DateTime CreatedAt { get; set; }

    [JsonProperty("updatedAt")]
    public DateTime UpdatedAt { get; set; }

    [JsonProperty("deletedAt")]
    public DateTime? DeletedAt { get; set; }

    [JsonProperty("responded")]
    public bool Responded { get; set; }

    [JsonProperty("respondBy")]
    public DateTime RespondBy { get; set; }

    [JsonProperty("feedbackRequested")]
    public bool FeedbackRequested { get; set; }

    [JsonProperty("reminderSent")]
    public bool ReminderSent { get; set; }

    [JsonProperty("priorityNote")]
    public string PriorityNote { get; set; }

    [JsonProperty("escalationInstructions")]
    public string EscalationInstructions { get; set; }

    [JsonProperty("testCase")]
    public bool TestCase { get; set; }

    [JsonProperty("customFields")]
    public List<CustomerData> CustomFields { get; set; }

    [JsonProperty("submittedBy")]
    public SubmittedBy SubmittedBy { get; set; }

    [JsonProperty("submitterContactDetails")]
    public SubmitterContactDetails SubmitterContactDetails { get; set; }

    [JsonProperty("caseNotes")]
    public List<CaseNote> CaseNotes { get; set; }

    [JsonProperty("caseResponses")]
    public List<CaseResponse> CaseResponses { get; set; }
}
