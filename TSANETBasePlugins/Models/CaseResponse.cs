using Newtonsoft.Json;
using System;

public class CaseResponse
{
    [JsonProperty("id")]
    public int Id { get; set; }

    [JsonProperty("type")]
    public string Type { get; set; }

    [JsonProperty("caseNumber")]
    public string CaseNumber { get; set; }

    [JsonProperty("engineerName")]
    public string EngineerName { get; set; }

    [JsonProperty("engineerPhone")]
    public string EngineerPhone { get; set; }

    [JsonProperty("engineerEmail")]
    public string EngineerEmail { get; set; }

    [JsonProperty("nextSteps")]
    public string NextSteps { get; set; }

    [JsonProperty("createdAt")]
    public DateTime CreatedAt { get; set; }
}
