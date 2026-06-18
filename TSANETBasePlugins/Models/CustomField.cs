using Newtonsoft.Json;
using System.Collections.Generic;

public class CustomField
{
    [JsonProperty("fieldId")]
    public int FieldId { get; set; }

    [JsonProperty("section")]
    public string Section { get; set; }

    [JsonProperty("label")]
    public string Label { get; set; }

    [JsonProperty("options")]
    public string Options { get; set; }

    [JsonProperty("additionalSettings")]
    public string AdditionalSettings { get; set; }

    [JsonProperty("type")]
    public string Type { get; set; }

    [JsonProperty("displayOrder")]
    public int DisplayOrder { get; set; }

    [JsonProperty("validationRules")]
    public string ValidationRules { get; set; }

    [JsonProperty("value")]
    public string Value { get; set; }

    [JsonProperty("selections")]
    public List<FieldSelection> Selections { get; set; }

    [JsonProperty("required")]
    public bool Required { get; set; }
}
