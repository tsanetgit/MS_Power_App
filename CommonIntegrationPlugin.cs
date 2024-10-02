using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Collections.Generic;

public class CommonIntegrationPlugin
{
    private readonly IOrganizationService _service;
    private readonly ITracingService _tracingService;
    private readonly string _apiUrl;
    private readonly string _clientId;
    private readonly string _clientSecret;

    public CommonIntegrationPlugin(IOrganizationService service, ITracingService tracingService)
    {
        _service = service;
        _tracingService = tracingService;
        _apiUrl = GetEnvVariable(_service, "ap_API_URL");
        _clientId = GetEnvVariable(_service, "ap_API_CLIENT_ID");
        _clientSecret = GetEnvVariable(_service, "ap_API_CLIENT_SECRET");
    }

    public static string GetEnvVariable(IOrganizationService service, string name)
    {
        var envVariables = new Dictionary<string, string>();
        string val = "";

        var query = new QueryExpression("environmentvariabledefinition")
        {
            Criteria = new FilterExpression
            {
                Conditions =
                    {
                        new ConditionExpression
                        {
                            AttributeName = "schemaname",
                            Operator = ConditionOperator.Equal,
                            Values = { name }
                        }
                    }
            },
            ColumnSet = new ColumnSet("statecode", "defaultvalue", "valueschema",
                                      "schemaname", "environmentvariabledefinitionid", "type"),
            LinkEntities =
            {
                new LinkEntity
                {
                    JoinOperator = JoinOperator.LeftOuter,
                    LinkFromEntityName = "environmentvariabledefinition",
                    LinkFromAttributeName = "environmentvariabledefinitionid",
                    LinkToEntityName = "environmentvariablevalue",
                    LinkToAttributeName = "environmentvariabledefinitionid",
                    Columns = new ColumnSet("statecode", "value", "environmentvariablevalueid"),
                    EntityAlias = "v"
                }
            }
        };

        var results = service.RetrieveMultiple(query);
        if (results?.Entities.Count > 0)
        {
            foreach (var entity in results.Entities)
            {
                var schemaName = entity.GetAttributeValue<string>("schemaname");
                var value = entity.GetAttributeValue<AliasedValue>("v.value")?.Value?.ToString();
                var defaultValue = entity.GetAttributeValue<string>("defaultvalue");

                if (schemaName != null && !envVariables.ContainsKey(schemaName))
                {
                    envVariables.Add(schemaName, string.IsNullOrEmpty(value) ? defaultValue : value);
                }
                val = string.IsNullOrEmpty(value) ? defaultValue : value;
            }
        }

        return val;
    }

    public async Task<string> Login()
    {
        try
        {
            _tracingService.Trace("Initiating login request.");

            using (HttpClient client = new HttpClient())
            {
                var requestBody = new
                {
                    username = _clientId,
                    password = _clientSecret
                };

                var json = JsonConvert.SerializeObject(requestBody);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                _tracingService.Trace("Sending login request to API.");
                var response = await client.PostAsync($"{_apiUrl}/login", content);

                if (!response.IsSuccessStatusCode)
                {
                    _tracingService.Trace("Login request failed.");
                    throw new InvalidOperationException("Failed to login.");
                }

                _tracingService.Trace("Login request succeeded.");
                var responseContent = await response.Content.ReadAsStringAsync();
                var tokenResponse = Newtonsoft.Json.JsonConvert.DeserializeObject<TokenResponse>(responseContent);

                _tracingService.Trace("Access token retrieved successfully.");
                return tokenResponse.access_token;
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in Login: {ex.Message}");
            throw;
        }
    }

    public async Task<string> GetCompanyByName(string companyName, string accessToken)
    {
        try
        {
            _tracingService.Trace($"Retrieving company details for: {companyName}");

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);

                var response = await client.GetAsync($"{_apiUrl}/companies/{companyName}");

                if (!response.IsSuccessStatusCode)
                {
                    _tracingService.Trace($"Failed to retrieve company details for '{companyName}'.");
                    throw new InvalidOperationException($"Failed to retrieve company details for '{companyName}'.");
                }

                _tracingService.Trace($"Company details retrieved successfully for: {companyName}");
                return await response.Content.ReadAsStringAsync();
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in GetCompanyByName: {ex.Message}");
            throw;
        }
    }

    public async Task<string> GetFormByCompany(int companyId, string accessToken)
    {
        try
        {
            _tracingService.Trace($"Retrieving form details for company ID: {companyId}");

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);

                var response = await client.GetAsync($"{_apiUrl}/form/company/{companyId}");

                if (!response.IsSuccessStatusCode)
                {
                    _tracingService.Trace($"Failed to retrieve form details for company ID '{companyId}'.");
                    throw new InvalidOperationException($"Failed to retrieve form details for company ID '{companyId}'.");
                }

                _tracingService.Trace($"Form details retrieved successfully for company ID: {companyId}");
                return await response.Content.ReadAsStringAsync();
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in GetFormByCompany: {ex.Message}");
            throw;
        }
    }


    public async Task<string> PostCase(object caseDetails, string accessToken)
    {
        try
        {
            _tracingService.Trace("Sending case details to API.");

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);

                var json = JsonConvert.SerializeObject(caseDetails);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await client.PostAsync($"{_apiUrl}/case", content);

                if (!response.IsSuccessStatusCode)
                {
                    _tracingService.Trace("Failed to create case.");
                    throw new InvalidOperationException("Failed to create case.");
                }

                _tracingService.Trace("Case created successfully.");
                return await response.Content.ReadAsStringAsync();
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in PostCase: {ex.Message}");
            throw;
        }
    }
}
public class TokenResponse
{
    public string access_token { get; set; }
    public string token_type { get; set; }
    public int expires_in { get; set; }
}
