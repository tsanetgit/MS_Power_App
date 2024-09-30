using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

public class CommonIntegrationPlugin
{
    private readonly IOrganizationService _service;
    private readonly string _apiUrl;
    private readonly string _clientId;
    private readonly string _clientSecret;

    public CommonIntegrationPlugin(IOrganizationService service)
    {
        _service = service;
        _apiUrl = GetEnvironmentVariable("API_URL");
        _clientId = GetEnvironmentVariable("API_CLIENT_ID");
        _clientSecret = GetEnvironmentVariable("API_CLIENT_SECRET");
    }
    //get env var
    private string GetEnvironmentVariable(string schemaName)
    {
        QueryExpression query = new QueryExpression("environmentvariablevalue")
        {
            ColumnSet = new ColumnSet("value")
        };
        query.Criteria.AddCondition("environmentvariabledefinitionid", ConditionOperator.Equal, schemaName);

        var result = _service.RetrieveMultiple(query);
        if (result.Entities.Count > 0)
        {
            return result.Entities[0].GetAttributeValue<string>("value");
        }
        throw new InvalidOperationException($"Environment variable '{schemaName}' not found.");
    }
    //main method to get 
    public async Task<string> Login()
    {
        using (HttpClient client = new HttpClient())
        {
            var requestBody = new
            {
                client_id = _clientId,
                client_secret = _clientSecret,
                grant_type = "client_credentials"
            };

            var json = Newtonsoft.Json.JsonConvert.SerializeObject(requestBody);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await client.PostAsync($"{_apiUrl}/identity/connect/token", content);

            if (!response.IsSuccessStatusCode)
            {
                throw new InvalidOperationException("Failed to login.");
            }

            var responseContent = await response.Content.ReadAsStringAsync();
            var tokenResponse = Newtonsoft.Json.JsonConvert.DeserializeObject<TokenResponse>(responseContent);

            return tokenResponse.access_token;
        }
    }
    //get company by name - fulltext
    public async Task<string> GetCompanyByName(string companyName, string accessToken)
    {
        using (HttpClient client = new HttpClient())
        {
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);

            var response = await client.GetAsync($"{_apiUrl}/companies/{companyName}");

            if (!response.IsSuccessStatusCode)
            {
                throw new InvalidOperationException($"Failed to retrieve company details for '{companyName}'.");
            }

            return await response.Content.ReadAsStringAsync();
        }
    }
    //get form by company id
    public async Task<string> GetFormByCompany(string companyId, string accessToken)
    {
        using (HttpClient client = new HttpClient())
        {
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);

            var response = await client.GetAsync($"{_apiUrl}/form/company/{companyId}");

            if (!response.IsSuccessStatusCode)
            {
                throw new InvalidOperationException($"Failed to retrieve form details for company ID '{companyId}'.");
            }

            return await response.Content.ReadAsStringAsync();
        }
    }

    // Method to send POST request to create a case
    public async Task<string> PostCase(object caseDetails, string accessToken)
    {
        using (HttpClient client = new HttpClient())
        {
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);

            var json = JsonConvert.SerializeObject(caseDetails);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await client.PostAsync($"{_apiUrl}/case", content);

            if (!response.IsSuccessStatusCode)
            {
                throw new InvalidOperationException("Failed to create case.");
            }

            return await response.Content.ReadAsStringAsync();
        }
    }
}

public class TokenResponse
{
    public string access_token { get; set; }
    public string token_type { get; set; }
    public int expires_in { get; set; }
}
