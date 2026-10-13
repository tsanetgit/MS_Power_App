using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace TSANETBasePlugins.E2ee
{
    internal static class E2eeUtil
    {
        public const string Profile = "tsanet-e2ee-1";

        /// <summary>Reads an ap_tsanetsetting row by ap_name, optionally also matching ap_value.</summary>
        public static Entity GetSetting(IOrganizationService service, string name, string value = null)
        {
            var query = new QueryExpression("ap_tsanetsetting")
            {
                ColumnSet = new ColumnSet("ap_value", "ap_uri"),
                TopCount = 1
            };
            query.Criteria.AddCondition("ap_name", ConditionOperator.Equal, name);
            if (value != null)
            {
                query.Criteria.AddCondition("ap_value", ConditionOperator.Equal, value);
            }
            EntityCollection result = service.RetrieveMultiple(query);
            return result.Entities.Count > 0 ? result.Entities[0] : null;
        }

        public static byte[] Sha256(byte[] data)
        {
            using (var sha = SHA256.Create())
            {
                return sha.ComputeHash(data);
            }
        }

        /// <summary>BASE64URL(SHA-256(UTF-8 bytes of value)).</summary>
        public static string Digest(string value)
        {
            return Base64UrlEncoding.Encode(Sha256(Encoding.UTF8.GetBytes(value)));
        }

        public static bool FixedTimeEquals(byte[] a, byte[] b)
        {
            if (a == null || b == null || a.Length != b.Length)
            {
                return false;
            }
            int diff = 0;
            for (int i = 0; i < a.Length; i++)
            {
                diff |= a[i] ^ b[i];
            }
            return diff == 0;
        }

        public static bool FixedTimeEquals(string a, string b)
        {
            return FixedTimeEquals(Encoding.ASCII.GetBytes(a ?? string.Empty), Encoding.ASCII.GetBytes(b ?? string.Empty));
        }

        /// <summary>Parses JSON without date/float coercion; throws on malformed input.</summary>
        public static JObject ParseObject(string json, string what)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                throw new InvalidPluginExecutionException(what + " is empty.");
            }
            try
            {
                using (var sr = new StringReader(json))
                using (var reader = new JsonTextReader(sr))
                {
                    reader.DateParseHandling = DateParseHandling.None;
                    reader.FloatParseHandling = FloatParseHandling.Decimal;
                    JToken token = JToken.ReadFrom(reader);
                    var obj = token as JObject;
                    if (obj == null)
                    {
                        throw new InvalidPluginExecutionException(what + " is not a JSON object.");
                    }
                    return obj;
                }
            }
            catch (JsonException ex)
            {
                throw new InvalidPluginExecutionException(what + " is not valid JSON: " + ex.Message, ex);
            }
        }

        public static string Serialize(JToken token)
        {
            return token.ToString(Formatting.None);
        }

        public static string StringOrNull(JToken token)
        {
            return token != null && token.Type == JTokenType.String ? (string)token : null;
        }

        public static void Trace(ITracingService tracing, string message)
        {
            if (tracing != null)
            {
                tracing.Trace("[E2EE] " + message);
            }
        }

        public static string RequireValue(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new InvalidPluginExecutionException(name + " is required.");
            }
            return value;
        }
    }
}
