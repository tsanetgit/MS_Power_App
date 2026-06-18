using Newtonsoft.Json;

public class SubmittedBy
{
    [JsonProperty("id")]
    public int Id { get; set; }

    [JsonProperty("username")]
    public string Username { get; set; }

    [JsonProperty("firstName")]
    public string FirstName { get; set; }

    [JsonProperty("lastName")]
    public string LastName { get; set; }

    [JsonProperty("email")]
    public string Email { get; set; }

    [JsonProperty("phone")]
    public string Phone { get; set; }

    [JsonProperty("phoneCountryCode")]
    public string PhoneCountryCode { get; set; }

    [JsonProperty("city")]
    public string City { get; set; }
}
