using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Net.Http.Headers;
using System.IO;
using System.IO.Compression;

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

    // Helper method to add default headers to the HttpClient
    private void AddDefaultHeaders(HttpClient client)
    {
        client.DefaultRequestHeaders.Add("Accept", "*/*");
        client.DefaultRequestHeaders.Add("Accept-Encoding", "gzip, deflate");
        client.DefaultRequestHeaders.Add("Connection", "keep-alive");
        client.DefaultRequestHeaders.Add("User-Agent", "Dynamics/9.1");
    }

    // Helper method to decompress response
    private async Task<string> DecompressResponse(HttpContent content, Stream responseStream)
    {
        string responseContent;

        if (content.Headers.ContentEncoding.Contains("gzip"))
        {
            _tracingService.Trace("Response is GZIP compressed.");
            using (var decompressedStream = new GZipStream(responseStream, CompressionMode.Decompress))
            using (var reader = new StreamReader(decompressedStream))
            {
                responseContent = await reader.ReadToEndAsync();
            }
        }
        else if (content.Headers.ContentEncoding.Contains("deflate"))
        {
            _tracingService.Trace("Response is Deflate compressed.");
            using (var decompressedStream = new DeflateStream(responseStream, CompressionMode.Decompress))
            using (var reader = new StreamReader(decompressedStream))
            {
                responseContent = await reader.ReadToEndAsync();
            }
        }
        else
        {
            _tracingService.Trace("Response is not compressed.");
            responseContent = await content.ReadAsStringAsync();
        }

        return responseContent;
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

                // Add default headers
                AddDefaultHeaders(client);
                content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                content.Headers.ContentLength = json.Length;

                _tracingService.Trace("Sending login request to API. " + _apiUrl);

                var response = await client.PostAsync($"{_apiUrl}/login", content);
                // Check if the response was successful
                if (!response.IsSuccessStatusCode)
                {
                    _tracingService.Trace("Login request failed. Response: " + await response.Content.ReadAsStringAsync());
                    throw new InvalidOperationException("Failed to login.");
                }

                // Check if the response was successful
                if (!response.IsSuccessStatusCode)
                {
                    _tracingService.Trace("Login request failed. Response: " + await response.Content.ReadAsStringAsync());
                    throw new InvalidOperationException("Failed to login.");
                }

                Stream responseStream = await response.Content.ReadAsStreamAsync();
                string responseContent = await DecompressResponse(response.Content, responseStream);

                // Deserialize the response
                var tokenResponse = JsonConvert.DeserializeObject<TokenResponse>(responseContent);

                if (tokenResponse == null || string.IsNullOrEmpty(tokenResponse.AccessToken))
                {
                    throw new InvalidOperationException("Invalid token response.");
                }

                _tracingService.Trace("Access token retrieved successfully.");
                return tokenResponse.AccessToken;
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
                // Add default headers
                AddDefaultHeaders(client);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                _tracingService.Trace("Sending request to get company details.");

                var response = await client.GetAsync($"{_apiUrl}/companies/{companyName}");

                // Check if the response was successful
                if (!response.IsSuccessStatusCode)
                {
                    _tracingService.Trace($"Failed to retrieve company details for '{companyName}'. Response: " + await response.Content.ReadAsStringAsync());
                    throw new InvalidOperationException($"Failed to retrieve company details for '{companyName}'.");
                }

                Stream responseStream = await response.Content.ReadAsStreamAsync();
                string responseContent = await DecompressResponse(response.Content, responseStream);

                _tracingService.Trace($"Company details for {companyName} retrieved successfully. Result: " + responseContent);
                return responseContent;
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in GetCompanyByName: {ex.Message}");
            throw;
        }
    }

    public async Task<ApiResponse> GetFormByCompany(int companyId, string accessToken)
    {
        var apiResponse = new ApiResponse();

        try
        {
            _tracingService.Trace($"Retrieving form details for company ID: {companyId}");

            using (HttpClient client = new HttpClient())
            {
                // Add default headers
                AddDefaultHeaders(client);
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);

                var response = await client.GetAsync($"{_apiUrl}/form/company/{companyId}");

                if (!response.IsSuccessStatusCode)
                {
                    _tracingService.Trace($"Failed to retrieve form details for company ID '{companyId}'. Status Code: {response.StatusCode}");

                    // Set the error in the response object
                    apiResponse.IsError = true;
                    apiResponse.Content = $"Error: Failed to retrieve form details for company ID '{companyId}' - Status Code: {response.StatusCode}, Message: {await response.Content.ReadAsStringAsync()}";
                    return apiResponse;
                }

                Stream responseStream = await response.Content.ReadAsStreamAsync();
                string responseContent = await DecompressResponse(response.Content, responseStream);

                _tracingService.Trace($"Form details retrieved successfully for company ID: {companyId}");

                // Set the success content in the response object
                apiResponse.IsError = false;
                apiResponse.Content = responseContent;
                return apiResponse;
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in GetFormByCompany: {ex.Message}");

            // Set the exception in the response object
            apiResponse.IsError = true;
            apiResponse.Content = $"Error: Exception occurred while retrieving form details for company ID '{companyId}' - {ex.Message}";
            return apiResponse;
        }
    }



    public async Task<string> PostCase(object caseDetails, string accessToken)
    {
        try
        {
            _tracingService.Trace("Sending case details to API.");

            using (HttpClient client = new HttpClient())
            {
                // Add default headers
                AddDefaultHeaders(client);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var json = JsonConvert.SerializeObject(caseDetails);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                _tracingService.Trace("Sending POST request to create case.");

                var response = await client.PostAsync($"{_apiUrl}/case", content);

                // Check if the response was successful
                if (!response.IsSuccessStatusCode)
                {
                    _tracingService.Trace("Failed to create case. Response: " + await response.Content.ReadAsStringAsync());
                    throw new InvalidOperationException("Failed to create case.");
                }

                Stream responseStream = await response.Content.ReadAsStreamAsync();
                string responseContent = await DecompressResponse(response.Content, responseStream);

                _tracingService.Trace("Case created successfully.");
                return responseContent;
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
    [JsonProperty("accessToken")]
    public string AccessToken { get; set; }

    [JsonProperty("tokenType")]
    public string TokenType { get; set; }

    [JsonProperty("expiresIn")]
    public int ExpiresIn { get; set; }
}

public class ApiResponse
{
    public bool IsError { get; set; }
    public string Content { get; set; }
}