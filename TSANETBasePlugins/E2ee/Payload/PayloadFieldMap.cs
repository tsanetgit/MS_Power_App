using Microsoft.Xrm.Sdk;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Globalization;

namespace TSANETBasePlugins.E2ee
{
    /// <summary>Shape of the payload being processed (custom-field id property differs).</summary>
    internal enum PayloadDirection
    {
        /// <summary>Payload we send (customFields items use "fieldId").</summary>
        Outgoing,
        /// <summary>Payload we receive (customFields items use "id").</summary>
        Incoming
    }

    internal sealed class ProtectedField
    {
        public string Id { get; set; }
        public JObject Container { get; set; }
        public string Property { get; set; }

        public string Value
        {
            get
            {
                string value = (string)Container[Property];
                if (value != null && value.StartsWith("{&#34;", System.StringComparison.Ordinal))
                {
                    value = System.Net.WebUtility.HtmlDecode(value);
                }
                return value;
            }
        }

        public void SetValue(string value)
        {
            Container[Property] = value;
        }
    }

    /// <summary>The single place to add protected fields.</summary>
    internal static class PayloadFieldMap
    {
        private static readonly string[] CollaborationRequestFields = { "problemSummary", "problemDescription" };
        private static readonly string[] NoteFields = { "summary", "description" };

        private const string CustomFieldsArray = "customFields";
        private const string CustomFieldValueProperty = "value";
        private const string OutgoingCustomFieldIdProperty = "fieldId";
        private const string IncomingCustomFieldIdProperty = "id";

        /// <summary>Returns the non-null protected fields present in the payload.</summary>
        public static List<ProtectedField> Enumerate(JObject payload, PayloadType type, PayloadDirection direction)
        {
            var result = new List<ProtectedField>();
            var seen = new HashSet<string>();

            string[] topLevel = type == PayloadType.Note ? NoteFields : CollaborationRequestFields;
            foreach (string name in topLevel)
            {
                string property = name;
                if (type == PayloadType.CollaborationRequest && direction == PayloadDirection.Incoming && IsNull(payload[name]))
                {
                    property = name == "problemSummary" ? "summary" : "description";
                }
                JToken token = payload[property];
                if (IsNull(token))
                {
                    continue;
                }
                RequireString(token, property);
                Add(result, seen, new ProtectedField { Id = name, Container = payload, Property = property });
            }

            if (type == PayloadType.CollaborationRequest)
            {
                JToken custom = payload[CustomFieldsArray];
                if (!IsNull(custom))
                {
                    var array = custom as JArray;
                    if (array == null)
                    {
                        throw new InvalidPluginExecutionException(CustomFieldsArray + " must be an array.");
                    }

                    string idProperty = direction == PayloadDirection.Outgoing ? OutgoingCustomFieldIdProperty : IncomingCustomFieldIdProperty;
                    foreach (JToken item in array)
                    {
                        var obj = item as JObject;
                        if (obj == null)
                        {
                            throw new InvalidPluginExecutionException(CustomFieldsArray + " items must be objects.");
                        }
                        JToken value = obj[CustomFieldValueProperty];
                        if (IsNull(value))
                        {
                            continue;
                        }
                        string id = ReadId(obj[idProperty], idProperty);
                        string fieldName = CustomFieldsArray + "/" + id + "/" + CustomFieldValueProperty;
                        RequireString(value, fieldName);
                        Add(result, seen, new ProtectedField { Id = fieldName, Container = obj, Property = CustomFieldValueProperty });
                    }
                }
            }

            return result;
        }

        private static void Add(List<ProtectedField> list, HashSet<string> seen, ProtectedField field)
        {
            if (!seen.Add(field.Id))
            {
                throw new InvalidPluginExecutionException("Duplicate protected field identifier '" + field.Id + "'.");
            }
            list.Add(field);
        }

        private static bool IsNull(JToken token)
        {
            return token == null || token.Type == JTokenType.Null || token.Type == JTokenType.Undefined;
        }

        private static void RequireString(JToken token, string name)
        {
            if (token.Type != JTokenType.String)
            {
                throw new InvalidPluginExecutionException("Protected field '" + name + "' must be a string.");
            }
        }

        private static string ReadId(JToken token, string property)
        {
            if (token == null || token.Type == JTokenType.Null)
            {
                throw new InvalidPluginExecutionException("Custom field item is missing '" + property + "'.");
            }
            if (token.Type == JTokenType.Integer || token.Type == JTokenType.String)
            {
                string id = System.Convert.ToString(((JValue)token).Value, CultureInfo.InvariantCulture);
                if (!string.IsNullOrEmpty(id))
                {
                    return id;
                }
            }
            throw new InvalidPluginExecutionException("Custom field '" + property + "' must be an integer or string.");
        }
    }
}
