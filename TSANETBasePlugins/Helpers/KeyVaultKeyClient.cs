using Microsoft.Xrm.Sdk;
using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

/// <summary>
/// Thin REST client for the Azure Key Vault "sign" and "unwrapKey" key operations,
/// authenticated via the Dataverse Managed Identity service. Private key material never
/// leaves Key Vault — only digests/wrapped keys are sent, and signatures/unwrapped keys
/// are returned.
/// </summary>
internal class KeyVaultKeyClient
{
    private const string KeyVaultApiVersion = "7.4";
    private static readonly string[] KeyVaultScopes = { "https://vault.azure.net/.default" };

    private readonly IManagedIdentityService _managedIdentityService;
    private readonly ITracingService _tracingService;

    public KeyVaultKeyClient(IServiceProvider serviceProvider)
    {
        _managedIdentityService = (IManagedIdentityService)serviceProvider.GetService(typeof(IManagedIdentityService));
        _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
    }

    /// <summary>
    /// Sends a SHA-256 digest to Key Vault's ES256 "sign" endpoint and returns the raw
    /// (base64url-decoded) signature bytes. <paramref name="signingKeyId"/> is the full
    /// Key Vault key identifier (e.g. https://{vault}.vault.azure.net/keys/{name}/{version}).
    /// </summary>
    public byte[] Sign(string signingKeyId, byte[] digest)
    {
        string endpoint = $"{signingKeyId.TrimEnd('/')}/sign?api-version={KeyVaultApiVersion}";
        var requestBody = new JObject
        {
            ["alg"] = "ES256",
            ["value"] = Base64UrlEncoding.Encode(digest)
        };

        JObject response = InvokeKeyVault(endpoint, requestBody);
        string signatureBase64Url = (string)response["value"];
        return Base64UrlEncoding.Decode(signatureBase64Url);
    }

    /// <summary>
    /// Sends an RSA-OAEP-256 wrapped CEK to Key Vault's "unwrapKey" endpoint and returns the
    /// recovered raw CEK bytes. <paramref name="unwrapKeyId"/> is the full Key Vault key
    /// identifier (e.g. https://{vault}.vault.azure.net/keys/{name}/{version}).
    /// </summary>
    public byte[] UnwrapKey(string unwrapKeyId, byte[] wrappedCek)
    {
        string endpoint = $"{unwrapKeyId.TrimEnd('/')}/unwrapkey?api-version={KeyVaultApiVersion}";
        var requestBody = new JObject
        {
            ["alg"] = "RSA-OAEP-256",
            ["value"] = Base64UrlEncoding.Encode(wrappedCek)
        };

        JObject response = InvokeKeyVault(endpoint, requestBody);
        string cekBase64Url = (string)response["value"];
        return Base64UrlEncoding.Decode(cekBase64Url);
    }

    /// <summary>
    /// Synchronously invokes <see cref="CallKeyVault"/> and surfaces the real underlying
    /// exception. Using <c>Task.Result</c> would wrap any exception thrown by the async
    /// method body (HTTP failures, managed identity token acquisition failures, JSON
    /// parsing errors, etc.) in an <see cref="AggregateException"/>, which the Dataverse
    /// sandbox host then reports back only as the generic "One or more errors occurred."
    /// message — hiding the actual root cause from the plugin trace log and the caller.
    /// <see cref="TaskAwaiter.GetResult"/> (via <c>GetAwaiter().GetResult()</c>) instead
    /// unwraps and rethrows the original exception, and the catch block below additionally
    /// traces the full exception details (including any inner exceptions) before rethrowing.
    /// </summary>
    private JObject InvokeKeyVault(string endpoint, JObject requestBody)
    {
        try
        {
            return CallKeyVault(endpoint, requestBody).GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            _tracingService?.Trace($"Key Vault call to {endpoint} failed: {ex}");

            if (ex is InvalidPluginExecutionException)
            {
                throw;
            }

            throw new InvalidPluginExecutionException($"Key Vault call to {endpoint} failed: {ex.Message}", ex);
        }
    }

    private async Task<JObject> CallKeyVault(string endpoint, JObject requestBody)
    {
        string accessToken = _managedIdentityService.AcquireToken(KeyVaultScopes);

        using (var client = new HttpClient())
        {
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            var content = new StringContent(requestBody.ToString(), Encoding.UTF8, "application/json");
            HttpResponseMessage response = await client.PostAsync(endpoint, content).ConfigureAwait(false);
            string responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

            if (!response.IsSuccessStatusCode)
            {
                _tracingService?.Trace($"Key Vault call to {endpoint} failed with status {response.StatusCode}: {responseContent}");
                throw new InvalidPluginExecutionException($"Key Vault operation failed with status {response.StatusCode}: {responseContent}");
            }

            return JObject.Parse(responseContent);
        }
    }
}
