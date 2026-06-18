using Newtonsoft.Json;

public class CustomerData
{
    [JsonProperty("id")]
    public int Id { get; set; }

    [JsonProperty("section")]
    public string Section { get; set; }

    [JsonProperty("fieldName")]
    public string FieldName { get; set; }

    [JsonProperty("value")]
    public string Value { get; set; }
}
