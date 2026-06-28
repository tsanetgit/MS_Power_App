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
using System.Linq;

public class CommonIntegrationPlugin
{
    private readonly IOrganizationService _service;
    private readonly ITracingService _tracingService;
    private readonly string _apiUrl;
    private readonly string _clientId;
    private readonly string _clientSecret;
    private readonly int _authorizationType;
    private readonly string _tenantId;
    private readonly string _oauth2Uri;
    private readonly string _oauth2Scope;

    public CommonIntegrationPlugin(IOrganizationService service, ITracingService tracingService)
    {
        _service = service;
        _tracingService = tracingService;

        var commonCasePlugin = new CommonCasePlugin();
        Entity settings = commonCasePlugin.GetIntegrationSettings(_service, _tracingService);
        _apiUrl = settings.GetAttributeValue<string>("ap_uri");
        _clientId = settings.GetAttributeValue<string>("ap_clientid");
        _clientSecret = settings.GetAttributeValue<string>("ap_secret");
        _tenantId = settings.GetAttributeValue<string>("ap_tenantid");

        var authTypeOptionSet = settings.GetAttributeValue<int>("ap_authorizationtype");
        _authorizationType = authTypeOptionSet;

        _oauth2Uri = GetEnvVariable(_service, "ap_OAauth2URI");
        _oauth2Scope = _clientId + "/.default";
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

    // Method to get environment variable value by name
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

    //Method to login based on the authorization type
    public async Task<string> Login()
    {
        if (_authorizationType == 1)
        {
            return await LoginOAuth2();
        }

        return await LoginLegacy();
    }

    // OAuth2 client credentials flow
    private async Task<string> LoginOAuth2()
    {
        try
        {
            _tracingService.Trace("Initiating OAuth2 client credentials login.");

            if (string.IsNullOrEmpty(_oauth2Uri))
                throw new InvalidOperationException("OAuth2 URI (ap_OAauth2URI) is not configured.");

            if (string.IsNullOrEmpty(_tenantId))
                throw new InvalidOperationException("Tenant ID (ap_tenantid) is not configured.");

            var tokenEndpoint = $"{_oauth2Uri.TrimEnd('/')}/{_tenantId}/oauth2/v2.0/token";
            _tracingService.Trace($"OAuth2 token endpoint: {tokenEndpoint}");

            using (HttpClient client = new HttpClient())
            {
                var formValues = new List<KeyValuePair<string, string>>
                {
                    new KeyValuePair<string, string>("grant_type", "client_credentials"),
                    new KeyValuePair<string, string>("client_id", _clientId),
                    new KeyValuePair<string, string>("client_secret", _clientSecret),
                    new KeyValuePair<string, string>("scope", _oauth2Scope)
                };

                var formContent = new FormUrlEncodedContent(formValues);

                _tracingService.Trace("Sending OAuth2 token request.");
                var response = await client.PostAsync(tokenEndpoint, formContent);

                if (!response.IsSuccessStatusCode)
                {
                    var errorBody = await response.Content.ReadAsStringAsync();
                    _tracingService.Trace($"OAuth2 token request failed. Status: {response.StatusCode}. Response: {errorBody}");
                    throw new InvalidOperationException($"OAuth2 login failed: {errorBody}");
                }

                var responseBody = await response.Content.ReadAsStringAsync();
                var tokenResponse = JsonConvert.DeserializeObject<OAuth2TokenResponse>(responseBody);

                if (tokenResponse == null || string.IsNullOrEmpty(tokenResponse.AccessToken))
                    throw new InvalidOperationException("OAuth2 token response was empty or invalid.");

                _tracingService.Trace("OAuth2 access token retrieved successfully.");
                return tokenResponse.AccessToken;
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in LoginOAuth2: {ex.Message}");
            throw;
        }
    }
    //Legacy login method for Basic Authentication
    public async Task<string> LoginLegacy()
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

                var response = await client.PostAsync($"{_apiUrl}/v1/login", content);
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

                var response = await client.GetAsync($"{_apiUrl}/v1/partners/{companyName}");

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
            _tracingService.Trace($"Starting GetFormByCompany for company ID: {companyId}");

            using (HttpClient client = new HttpClient())
            {
                // Add default headers
                AddDefaultHeaders(client);
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);

                _tracingService.Trace("Sending request to retrieve form details.");
                var response = await client.GetAsync($"{_apiUrl}/v1/forms/company/{companyId}");

                if (!response.IsSuccessStatusCode)
                {
                    _tracingService.Trace($"Failed to retrieve form details for company ID '{companyId}'. Status Code: {response.StatusCode}");
                    apiResponse.IsError = true;
                    apiResponse.Content = await response.Content.ReadAsStringAsync();
                    return apiResponse;
                }

                Stream responseStream = await response.Content.ReadAsStreamAsync();
                string responseContent = await DecompressResponse(response.Content, responseStream);

                _tracingService.Trace("Form details received, starting deserialization.");
                //var formData = JsonConvert.DeserializeObject<FormResponse>(responseContent);

                //if (formData == null || formData.CustomerData == null)
                //{
                //    _tracingService.Trace("Error: formData or formData.CustomerData is null.");
                //    apiResponse.IsError = true;
                //    apiResponse.Content = "Error: form data or customer data is null.";
                //    return apiResponse;
                //}

                //_tracingService.Trace("Processing customerData objects.");

                //// Process each customerData object to retrieve field metadata and possibly field selections
                //foreach (var customerData in formData.CustomerData)
                //{
                //    _tracingService.Trace($"Processing customerData ID: {customerData.Id}");

                //    // Get field metadata
                //    var fieldMetadata = await GetFieldMetadata(customerData.Id, formData.DocumentId, accessToken);

                //    if (fieldMetadata == null)
                //    {
                //        _tracingService.Trace($"Warning: Field metadata for customerData ID '{customerData.Id}' is null.");
                //        continue;
                //    }

                //    // Add fieldMetadata to customerData
                //    customerData.FieldMetadata = fieldMetadata;

                //    // If the field type is TIERSELECT, get field selections
                //    if (fieldMetadata.Type == "TIERSELECT")
                //    {
                //        _tracingService.Trace($"Field type is TIERSELECT for customerData ID: {customerData.Id}, retrieving field selections.");
                //        var fieldSelections = await GetFieldSelections(customerData.Id, formData.DocumentId, accessToken);

                //        if (fieldSelections == null)
                //        {
                //            _tracingService.Trace($"Warning: Field selections for customerData ID '{customerData.Id}' is null.");
                //            continue;
                //        }

                //        customerData.FieldSelections = fieldSelections;
                //    }
                //}

                _tracingService.Trace($"Form details successfully processed for company ID: {companyId}");

                // Set the success content in the response object
                apiResponse.IsError = false;
                apiResponse.Content = responseContent;
                return apiResponse;
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in GetFormByCompany: {ex.Message}");
            apiResponse.IsError = true;
            apiResponse.Content = $"Error: Exception occurred while retrieving form details for company ID '{companyId}' - {ex.Message}";
            return apiResponse;
        }
    }

    public async Task<ApiResponse> GetFormByDepartment(int departmentId, string accessToken)
    {
        var apiResponse = new ApiResponse();

        try
        {
            _tracingService.Trace($"Starting GetFormByDepartment for department ID: {departmentId}");

            using (HttpClient client = new HttpClient())
            {
                // Add default headers
                AddDefaultHeaders(client);
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);

                _tracingService.Trace("Sending request to retrieve form details.");
                var response = await client.GetAsync($"{_apiUrl}/v1/forms/department/{departmentId}");

                if (!response.IsSuccessStatusCode)
                {
                    _tracingService.Trace($"Failed to retrieve form details for department ID '{departmentId}'. Status Code: {response.StatusCode}");
                    apiResponse.IsError = true;
                    apiResponse.Content = await response.Content.ReadAsStringAsync();
                    return apiResponse;
                }

                Stream responseStream = await response.Content.ReadAsStreamAsync();
                string responseContent = await DecompressResponse(response.Content, responseStream);

                _tracingService.Trace("Form details received, starting deserialization.");

                _tracingService.Trace($"Form details successfully processed for department ID: {departmentId}");

                // Set the success content in the response object
                apiResponse.IsError = false;
                apiResponse.Content = responseContent;
                return apiResponse;
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in GetFormByDepartment: {ex.Message}");
            apiResponse.IsError = true;
            apiResponse.Content = $"Error: Exception occurred while retrieving form details for department ID '{departmentId}' - {ex.Message}";
            return apiResponse;
        }
    }

    // Get user information from the /me endpoint
    public async Task<ApiResponse> GetMe(string accessToken)
    {
        var apiResponse = new ApiResponse();

        try
        {
            _tracingService.Trace("Retrieving user information from /me endpoint");

            using (HttpClient client = new HttpClient())
            {
                // Add default headers
                AddDefaultHeaders(client);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                _tracingService.Trace("Sending request to get user information.");

                var response = await client.GetAsync($"{_apiUrl}/v1/me");

                // Check if the response was successful
                if (!response.IsSuccessStatusCode)
                {
                    string resp = await response.Content.ReadAsStringAsync();
                    _tracingService.Trace($"Failed to retrieve user information. Response: " + resp);
                    apiResponse.IsError = true;
                    apiResponse.Content = resp;
                    return apiResponse;
                }

                Stream responseStream = await response.Content.ReadAsStreamAsync();
                string responseContent = await DecompressResponse(response.Content, responseStream);

                _tracingService.Trace("User information retrieved successfully.");
                apiResponse.IsError = false;
                apiResponse.Content = responseContent;
                return apiResponse;
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in GetMe: {ex.Message}");
            apiResponse.IsError = true;
            apiResponse.Content = $"Exception: {ex.Message}";
            return apiResponse;
        }
    }

    //DEPRECATED
    //public async Task<FieldMetadata> GetFieldMetadata(int customerDataId, int documentId, string accessToken)
    //{
    //    _tracingService.Trace($"Retrieving field metadata for customerDataId: {customerDataId}, documentId: {documentId}");

    //    try
    //    {
    //        using (HttpClient client = new HttpClient())
    //        {
    //            // Add default headers
    //            AddDefaultHeaders(client);
    //            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);

    //            var response = await client.GetAsync($"{_apiUrl}/0.1.1/form/{documentId}/field/{customerDataId}");

    //            if (!response.IsSuccessStatusCode)
    //            {
    //                _tracingService.Trace($"Failed to retrieve field metadata for customerDataId '{customerDataId}', documentId '{documentId}'. Status Code: {response.StatusCode}");
    //                return null; // Handle the failure accordingly
    //            }

    //            Stream responseStream = await response.Content.ReadAsStreamAsync();
    //            string responseContent = await DecompressResponse(response.Content, responseStream);

    //            _tracingService.Trace($"Field metadata retrieved successfully for customerDataId: {customerDataId}, documentId: {documentId}");
    //            return JsonConvert.DeserializeObject<FieldMetadata>(responseContent);
    //        }
    //    }
    //    catch (Exception ex)
    //    {
    //        _tracingService.Trace($"Exception in GetFieldMetadata: {ex.Message}");
    //        return null;
    //    }
    //}


    //public async Task<List<FieldSelection>> GetFieldSelections(int customerDataId, int documentId, string accessToken)
    //{
    //    _tracingService.Trace($"Retrieving field selections for customerDataId: {customerDataId}, documentId: {documentId}");

    //    try
    //    {
    //        using (HttpClient client = new HttpClient())
    //        {
    //            // Add default headers
    //            AddDefaultHeaders(client);
    //            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);

    //            var response = await client.GetAsync($"{_apiUrl}/0.1.1/form/{documentId}/field/{customerDataId}/selections");

    //            if (!response.IsSuccessStatusCode)
    //            {
    //                _tracingService.Trace($"Failed to retrieve field selections for customerDataId '{customerDataId}', documentId '{documentId}'. Status Code: {response.StatusCode}");
    //                return null; // Handle the failure accordingly
    //            }

    //            Stream responseStream = await response.Content.ReadAsStreamAsync();
    //            string responseContent = await DecompressResponse(response.Content, responseStream);

    //            _tracingService.Trace($"Field selections retrieved successfully for customerDataId: {customerDataId}, documentId: {documentId}");
    //            return JsonConvert.DeserializeObject<List<FieldSelection>>(responseContent);
    //        }
    //    }
    //    catch (Exception ex)
    //    {
    //        _tracingService.Trace($"Exception in GetFieldSelections: {ex.Message}");
    //        return null;
    //    }
    //}


    public async Task<ApiResponse> PostCase(string caseFormString, string accessToken)
    {
        var apiResponse = new ApiResponse();

        try
        {
            _tracingService.Trace("Sending case details to API.");

            using (HttpClient client = new HttpClient())
            {
                _tracingService.Trace("Deserializing caseDetails to object.");
                var caseDetails = JsonConvert.DeserializeObject<CaseForm>(caseFormString);

                // Add default headers
                AddDefaultHeaders(client);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var json = JsonConvert.SerializeObject(caseDetails);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                //content.Headers.ContentLength = json.Length;

                _tracingService.Trace("Sending POST request to create case.");

                var response = await client.PostAsync($"{_apiUrl}/v1/collaboration-requests", content);

                // Check if the response was successful
                if (!response.IsSuccessStatusCode)
                {
                    string resp = await response.Content.ReadAsStringAsync();
                    _tracingService.Trace("Failed to create case. Response: " + resp);
                    apiResponse.IsError = true;
                    apiResponse.Content = resp;
                    return apiResponse;
                }

                Stream responseStream = await response.Content.ReadAsStreamAsync();
                string responseContent = await DecompressResponse(response.Content, responseStream);

                _tracingService.Trace("Case created successfully.");
                apiResponse.IsError = false;
                apiResponse.Content = responseContent;
                return apiResponse;
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in PostCase: {ex.Message}");
            apiResponse.IsError = true;
            apiResponse.Content = $"Exception: {ex.Message}";
            return apiResponse;
        }
    }
    public async Task<ApiResponse> GetCase(string caseToken, string accessToken)
    {
        var apiResponse = new ApiResponse();

        try
        {
            _tracingService.Trace($"Retrieving case update for internal case number: {caseToken}");

            using (HttpClient client = new HttpClient())
            {
                // Add default headers
                AddDefaultHeaders(client);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                _tracingService.Trace("Sending request to get case update.");

                var response = await client.GetAsync($"{_apiUrl}/v1/collaboration-requests/{caseToken}");

                // Check if the response was successful
                if (!response.IsSuccessStatusCode)
                {
                    string resp = await response.Content.ReadAsStringAsync();
                    _tracingService.Trace($"Failed to retrieve case update for internal case number '{caseToken}'. Response: " + resp);
                    apiResponse.IsError = true;
                    apiResponse.Content = resp;
                    return apiResponse;
                }

                Stream responseStream = await response.Content.ReadAsStreamAsync();
                string responseContent = await DecompressResponse(response.Content, responseStream);

                _tracingService.Trace("Raw JSON response content: " + responseContent);

                _tracingService.Trace("Deserializing response content to list of Case objects.");
                var caseJson = JsonConvert.DeserializeObject<Case>(responseContent);

                if (caseJson == null)
                {
                    _tracingService.Trace("No cases found in the response.");
                    apiResponse.IsError = true;
                    apiResponse.Content = "No cases found.";
                    return apiResponse;
                }

                _tracingService.Trace($"Case update for case token {caseToken} retrieved successfully. Result: " + JsonConvert.SerializeObject(caseJson));
                apiResponse.IsError = false;
                apiResponse.Content = JsonConvert.SerializeObject(caseJson);
                return apiResponse;
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in GetCase: {ex.Message}");
            apiResponse.IsError = true;
            apiResponse.Content = $"Exception: {ex.Message}";
            return apiResponse;
        }
    }
    public async Task<ApiResponse> PostCaseNote(string caseToken, string summary, string description, string priority, SubmittedBy submittedBy, string accessToken)
    {
        var apiResponse = new ApiResponse();

        try
        {
            _tracingService.Trace("Sending case note details to API.");

            using (HttpClient client = new HttpClient())
            {
                var noteDetails = new
                {
                    summary = summary,
                    description = description,
                    priority = priority,
                    submittedBy = submittedBy
                };

                // Add default headers
                AddDefaultHeaders(client);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var json = JsonConvert.SerializeObject(noteDetails);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                //content.Headers.ContentLength = json.Length;

                _tracingService.Trace("Sending POST request to create case note.");

                var response = await client.PostAsync($"{_apiUrl}/v1/collaboration-requests/{caseToken}/notes", content);

                // Check if the response was successful
                if (!response.IsSuccessStatusCode)
                {
                    string resp = await response.Content.ReadAsStringAsync();
                    _tracingService.Trace("Failed to create case note. Response: " + resp);
                    apiResponse.IsError = true;
                    apiResponse.Content = resp;
                    return apiResponse;
                }

                // Deserialize the response content into a Case object
                Stream responseStream = await response.Content.ReadAsStreamAsync();
                string responseContent = await DecompressResponse(response.Content, responseStream);

                _tracingService.Trace("Case note created successfully. /n" + responseContent);
                apiResponse.IsError = false;
                apiResponse.Content = responseContent;
                return apiResponse;
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in PostCaseNote: {ex.Message}");
            apiResponse.IsError = true;
            apiResponse.Content = $"Exception: {ex.Message}";
            return apiResponse;
        }
    }
    public async Task<ApiResponse> PostCaseApproval(string caseToken, string caseNumber, string engineerName, string engineerPhone, string engineerEmail, string nextSteps, string accessToken)
    {
        var apiResponse = new ApiResponse();

        try
        {
            _tracingService.Trace("Sending case approval update details to API.");

            using (HttpClient client = new HttpClient())
            {
                var details = new
                {
                    caseNumber = caseNumber,
                    engineerName = engineerName,
                    engineerPhone = engineerPhone,
                    engineerEmail = engineerEmail,
                    nextSteps = nextSteps
                };

                // Add default headers
                AddDefaultHeaders(client);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var json = JsonConvert.SerializeObject(details);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                //content.Headers.ContentLength = json.Length;

                _tracingService.Trace("Sending POST request to post case approval. " + json);

                var response = await client.PostAsync($"{_apiUrl}/v1/collaboration-requests/{caseToken}/approval", content);

                // Check if the response was successful
                if (!response.IsSuccessStatusCode)
                {
                    string resp = await response.Content.ReadAsStringAsync();
                    _tracingService.Trace("Failed to post case approval. Response: " + resp);
                    apiResponse.IsError = true;
                    apiResponse.Content = resp;
                    return apiResponse;
                }

                Stream responseStream = await response.Content.ReadAsStreamAsync();
                string responseContent = await DecompressResponse(response.Content, responseStream);

                _tracingService.Trace("Case approval created successfully.");
                apiResponse.IsError = false;
                apiResponse.Content = responseContent;
                return apiResponse;
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in Case Approval: {ex.Message}");
            apiResponse.IsError = true;
            apiResponse.Content = $"Exception: {ex.Message}";
            return apiResponse;
        }
    }
    public async Task<ApiResponse> PostCaseInformationResponse(string caseToken, string engineerName, string engineerPhone, string engineerEmail, string requestedInformation, string accessToken)
    {
        var apiResponse = new ApiResponse();

        try
        {
            _tracingService.Trace("Sending case information details to API.");

            using (HttpClient client = new HttpClient())
            {
                var details = new
                {
                    requestedInformation = requestedInformation
                };

                // Add default headers
                AddDefaultHeaders(client);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var json = JsonConvert.SerializeObject(details);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                //content.Headers.ContentLength = json.Length;

                _tracingService.Trace("Sending POST request to post case information response information.");

                var response = await client.PostAsync($"{_apiUrl}/v1/collaboration-requests/{caseToken}/information-response", content);

                // Check if the response was successful
                if (!response.IsSuccessStatusCode)
                {
                    string resp = await response.Content.ReadAsStringAsync();
                    _tracingService.Trace("Failed to post case response information. Response: " + resp);
                    apiResponse.IsError = true;
                    apiResponse.Content = resp;
                    return apiResponse;
                }

                Stream responseStream = await response.Content.ReadAsStreamAsync();
                string responseContent = await DecompressResponse(response.Content, responseStream);

                _tracingService.Trace("Case information response created successfully.");
                apiResponse.IsError = false;
                apiResponse.Content = responseContent;
                return apiResponse;
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in Case response information: {ex.Message}");
            apiResponse.IsError = true;
            apiResponse.Content = $"Exception: {ex.Message}";
            return apiResponse;
        }
    }

    public async Task<ApiResponse> PostCaseRequestInformation(string caseToken, string engineerName, string engineerPhone, string engineerEmail, string requestedInformation, string accessToken)
    {
        var apiResponse = new ApiResponse();

        try
        {
            _tracingService.Trace("Sending case information details to API.");

            using (HttpClient client = new HttpClient())
            {
                var details = new
                {
                    engineerName = engineerName,
                    engineerPhone = engineerPhone,
                    engineerEmail = engineerEmail,
                    requestedInformation = requestedInformation
                };

                // Add default headers
                AddDefaultHeaders(client);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var json = JsonConvert.SerializeObject(details);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                //content.Headers.ContentLength = json.Length;

                _tracingService.Trace("Sending POST request to post case requested information.");

                var response = await client.PostAsync($"{_apiUrl}/v1/collaboration-requests/{caseToken}/information-request", content);

                // Check if the response was successful
                if (!response.IsSuccessStatusCode)
                {
                    string resp = await response.Content.ReadAsStringAsync();
                    _tracingService.Trace("Failed to post case request information. Response: " + resp);
                    apiResponse.IsError = true;
                    apiResponse.Content = resp;
                    return apiResponse;
                }

                Stream responseStream = await response.Content.ReadAsStreamAsync();
                string responseContent = await DecompressResponse(response.Content, responseStream);

                _tracingService.Trace("Case information request created successfully.");
                apiResponse.IsError = false;
                apiResponse.Content = responseContent;
                return apiResponse;
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in Case request information: {ex.Message}");
            apiResponse.IsError = true;
            apiResponse.Content = $"Exception: {ex.Message}";
            return apiResponse;
        }
    }

    public async Task<ApiResponse> PostCaseReject(string caseToken, string engineerName, string engineerPhone, string engineerEmail, string reason, string accessToken)
    {
        var apiResponse = new ApiResponse();

        try
        {
            _tracingService.Trace("Sending case reject update details to API.");

            using (HttpClient client = new HttpClient())
            {
                var details = new
                {
                    engineerName = engineerName,
                    engineerPhone = engineerPhone,
                    engineerEmail = engineerEmail,
                    reason = reason
                };

                // Add default headers
                AddDefaultHeaders(client);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var json = JsonConvert.SerializeObject(details);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                //content.Headers.ContentLength = json.Length;

                _tracingService.Trace("Sending POST request to post case reject.");

                var response = await client.PostAsync($"{_apiUrl}/v1/collaboration-requests/{caseToken}/rejection", content);

                // Check if the response was successful
                if (!response.IsSuccessStatusCode)
                {
                    string resp = await response.Content.ReadAsStringAsync();
                    _tracingService.Trace("Failed to post case reject. Response: " + resp);
                    apiResponse.IsError = true;
                    apiResponse.Content = resp;
                    return apiResponse;
                }

                Stream responseStream = await response.Content.ReadAsStreamAsync();
                string responseContent = await DecompressResponse(response.Content, responseStream);

                _tracingService.Trace("Case reject created successfully.");
                apiResponse.IsError = false;
                apiResponse.Content = responseContent;
                return apiResponse;
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in Post Case Reject: {ex.Message}");
            apiResponse.IsError = true;
            apiResponse.Content = $"Exception: {ex.Message}";
            return apiResponse;
        }
    }

    public async Task<ApiResponse> PostCaseClose(string caseToken, string accessToken)
    {
        var apiResponse = new ApiResponse();

        try
        {
            _tracingService.Trace("Sending case close update details to API.");

            using (HttpClient client = new HttpClient())
            {
                var details = new
                {
                };

                // Add default headers
                AddDefaultHeaders(client);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var json = JsonConvert.SerializeObject(details);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                //content.Headers.ContentLength = json.Length;

                _tracingService.Trace("Sending POST request to post case close.");

                var response = await client.PostAsync($"{_apiUrl}/v1/collaboration-requests/{caseToken}/closure", content);

                // Check if the response was successful
                if (!response.IsSuccessStatusCode)
                {
                    string resp = await response.Content.ReadAsStringAsync();
                    _tracingService.Trace("Failed to post case close. Response: " + resp);
                    apiResponse.IsError = true;
                    apiResponse.Content = resp;
                    return apiResponse;
                }

                Stream responseStream = await response.Content.ReadAsStreamAsync();
                string responseContent = await DecompressResponse(response.Content, responseStream);

                _tracingService.Trace("Case reject close successfully.");
                apiResponse.IsError = false;
                apiResponse.Content = responseContent;
                return apiResponse;
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in Post Case Reject: {ex.Message}");
            apiResponse.IsError = true;
            apiResponse.Content = $"Exception: {ex.Message}";
            return apiResponse;
        }
    }
    //DEPRECATED
    public async Task<ApiResponse> UpdateCaseApproval(int caseId, string caseNumber, string engineerName, string engineerPhone, string engineerEmail, string nextSteps, string accessToken)
    {
        var apiResponse = new ApiResponse();

        try
        {
            _tracingService.Trace("Sending case approval update details to API.");

            using (HttpClient client = new HttpClient())
            {
                var approvalDetails = new
                {
                    caseNumber = caseNumber,
                    engineerName = engineerName,
                    engineerPhone = engineerPhone,
                    engineerEmail = engineerEmail,
                    nextSteps = nextSteps
                };

                // Add default headers
                AddDefaultHeaders(client);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var json = JsonConvert.SerializeObject(approvalDetails);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                //content.Headers.ContentLength = json.Length;

                _tracingService.Trace("Sending POST request to update case approval.");

                var response = await client.PostAsync($"{_apiUrl}/0.1.1/cases/{caseId}/update/approval", content);

                // Check if the response was successful
                if (!response.IsSuccessStatusCode)
                {
                    string resp = await response.Content.ReadAsStringAsync();
                    _tracingService.Trace("Failed to update case approval. Response: " + resp);
                    apiResponse.IsError = true;
                    apiResponse.Content = resp;
                    return apiResponse;
                }

                Stream responseStream = await response.Content.ReadAsStreamAsync();
                string responseContent = await DecompressResponse(response.Content, responseStream);

                _tracingService.Trace("Case approval updated successfully.");
                apiResponse.IsError = false;
                apiResponse.Content = responseContent;
                return apiResponse;
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in UpdateCaseApproval: {ex.Message}");
            apiResponse.IsError = true;
            apiResponse.Content = $"Exception: {ex.Message}";
            return apiResponse;
        }
    }

    public async Task<ApiResponse> PatchCaseApproval(string caseToken, string caseNumber, string engineerName, string engineerPhone, string engineerEmail, string accessToken)
    {
        var apiResponse = new ApiResponse();

        try
        {
            _tracingService.Trace("Sending case approval update details to API using PATCH.");

            using (HttpClient client = new HttpClient())
            {
                var approvalDetails = new
                {
                    caseNumber = caseNumber,
                    engineerName = engineerName,
                    engineerPhone = engineerPhone,
                    engineerEmail = engineerEmail
                };

                // Add default headers
                AddDefaultHeaders(client);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var json = JsonConvert.SerializeObject(approvalDetails);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                //content.Headers.ContentLength = json.Length;

                _tracingService.Trace("Sending PATCH request to update case approval.");

                var request = new HttpRequestMessage(new HttpMethod("PATCH"), $"{_apiUrl}/v1/collaboration-requests/{caseToken}/approval")
                {
                    Content = content
                };

                var response = await client.SendAsync(request);

                // Check if the response was successful
                if (!response.IsSuccessStatusCode)
                {
                    string resp = await response.Content.ReadAsStringAsync();
                    _tracingService.Trace("Failed to update case approval. Response: " + resp);
                    apiResponse.IsError = true;
                    apiResponse.Content = resp;
                    return apiResponse;
                }

                Stream responseStream = await response.Content.ReadAsStreamAsync();
                string responseContent = await DecompressResponse(response.Content, responseStream);

                _tracingService.Trace("Case approval updated successfully using PATCH.");
                apiResponse.IsError = false;
                apiResponse.Content = responseContent;
                return apiResponse;
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in PatchCaseApproval: {ex.Message}");
            apiResponse.IsError = true;
            apiResponse.Content = $"Exception: {ex.Message}";
            return apiResponse;
        }
    }

    public async Task<ApiResponse> PatchCollaborationRequest(string token, string internalCaseNumber, SubmitterContactDetails submitterContactDetails, string accessToken)
    {
        var apiResponse = new ApiResponse();

        try
        {
            _tracingService.Trace($"Updating collaboration request for token: {token}");

            using (HttpClient client = new HttpClient())
            {
                var requestDetails = new
                {
                    internalCaseNumber = internalCaseNumber,
                    submitterContactDetails = submitterContactDetails
                };

                // Add default headers
                AddDefaultHeaders(client);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var json = JsonConvert.SerializeObject(requestDetails);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                //content.Headers.ContentLength = json.Length;

                _tracingService.Trace("Sending PATCH request to update collaboration request.");

                var request = new HttpRequestMessage(new HttpMethod("PATCH"), $"{_apiUrl}/v1/collaboration-requests/{token}")
                {
                    Content = content
                };

                var response = await client.SendAsync(request);

                // Check if the response was successful
                if (!response.IsSuccessStatusCode)
                {
                    string resp = await response.Content.ReadAsStringAsync();
                    _tracingService.Trace($"Failed to update collaboration request for token '{token}'. Response: " + resp);
                    apiResponse.IsError = true;
                    apiResponse.Content = resp;
                    return apiResponse;
                }

                Stream responseStream = await response.Content.ReadAsStreamAsync();
                string responseContent = await DecompressResponse(response.Content, responseStream);

                _tracingService.Trace($"Collaboration request for token {token} updated successfully.");
                apiResponse.IsError = false;
                apiResponse.Content = responseContent;
                return apiResponse;
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in PatchCollaborationRequest: {ex.Message}");
            apiResponse.IsError = true;
            apiResponse.Content = $"Exception: {ex.Message}";
            return apiResponse;
        }
    }

    // DEPRECATED
    public async Task<ApiResponse> GetAttachmentConfig(string token, string accessToken)
    {
        var apiResponse = new ApiResponse();

        try
        {
            _tracingService.Trace($"Retrieving attachment config for token: {token}");

            using (HttpClient client = new HttpClient())
            {
                // Add default headers
                AddDefaultHeaders(client);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                _tracingService.Trace("Sending request to get attachment config.");

                var response = await client.GetAsync($"{_apiUrl}/v1/collaboration-requests/{token}/attachments/config");

                // Check if the response was successful
                if (!response.IsSuccessStatusCode)
                {
                    string resp = await response.Content.ReadAsStringAsync();
                    _tracingService.Trace($"Failed to retrieve attachment config for token '{token}'. Response: " + resp);
                    apiResponse.IsError = true;
                    apiResponse.Content = resp;
                    return apiResponse;
                }

                Stream responseStream = await response.Content.ReadAsStreamAsync();
                string responseContent = await DecompressResponse(response.Content, responseStream);

                _tracingService.Trace("Attachment config retrieved successfully.");
                apiResponse.IsError = false;
                apiResponse.Content = responseContent;
                return apiResponse;
            }
        }
        catch (Exception ex)
        {
            _tracingService.Trace($"Exception in GetAttachmentConfig: {ex.Message}");
            apiResponse.IsError = true;
            apiResponse.Content = $"Exception: {ex.Message}";
            return apiResponse;
        }
    }
}

