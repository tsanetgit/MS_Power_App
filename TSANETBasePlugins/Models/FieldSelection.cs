using Newtonsoft.Json;
using System.Collections.Generic;

public class FieldSelection
{
    [JsonProperty("value")]
    public string Value { get; set; }

    [JsonProperty("children")]
    public List<FieldSelection> Children { get; set; }
}
