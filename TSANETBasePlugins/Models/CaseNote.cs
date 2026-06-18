using Newtonsoft.Json;
using System;

public class CaseNote
{
    [JsonProperty("id")]
    public int Id { get; set; }

    [JsonProperty("caseId")]
    public int CaseId { get; set; }

    [JsonProperty("companyName")]
    public string CompanyName { get; set; }

    [JsonProperty("creatorUsername")]
    public string CreatorUsername { get; set; }

    [JsonProperty("creatorEmail")]
    public string CreatorEmail { get; set; }

    [JsonProperty("creatorName")]
    public string CreatorName { get; set; }

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
}
