using Newtonsoft.Json;

public class MeResponse
{
    [JsonProperty("company")]
    public MeCompany Company { get; set; }

    [JsonProperty("user")]
    public MeUser User { get; set; }
}

public class MeCompany
{
    [JsonProperty("id")]
    public int? Id { get; set; }

    [JsonProperty("name")]
    public string Name { get; set; }

    [JsonProperty("description")]
    public string Description { get; set; }

    [JsonProperty("domain")]
    public string Domain { get; set; }
}

public class MeUser
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

    [JsonProperty("emailCheckFailData")]
    public string EmailCheckFailData { get; set; }
}
