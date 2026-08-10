using System;
using System.Collections.Generic;
using System.Text.Json;

namespace PnP.Core.Provisioning.ObjectHandlers.Utilities
{
    /// <summary>
    /// Reads the verbose OData shapes SharePoint's REST endpoints answer with.
    /// </summary>
    /// <remarks>
    /// <para><b>PnP Core talks to SharePoint REST in verbose OData</b> - <c>BatchClient</c> sends
    /// <c>Content-Type: application/json;odata=verbose</c> - and the responses come back in kind:</para>
    /// <list type="bullet">
    /// <item><description>a result is wrapped as <c>{"d":{"MethodName":…}}</c></description></item>
    /// <item><description>a collection is <c>{"results":[…]}</c>, not a bare array</description></item>
    /// </list>
    /// <para><b>This exists because the same mistake was made twice.</b> Both
    /// <see cref="SiteScriptUtility"/> and <see cref="StorageEntities"/> were written to read
    /// <c>value</c>, which is what the documentation's examples show and what the <em>nometadata</em>
    /// format uses. Reading <c>value</c> against a verbose response finds nothing and returns
    /// null or an empty list - so a write succeeds, the read that follows reports the thing does not
    /// exist, and nothing errors anywhere. Fixing it in one place and leaving the other cost a second
    /// live run to rediscover.</para>
    /// </remarks>
    internal static class VerboseOData
    {
        /// <summary>
        /// Strips the <c>d</c> / method-name / <c>results</c> envelope off a response.
        /// </summary>
        /// <remarks>
        /// <para>Unwrapped by shape rather than by method name, so one helper covers every endpoint
        /// without a table of names to keep in step. A response that is already unwrapped passes
        /// through.</para>
        /// <para><b><c>d</c> is used two different ways</b>, and telling them apart is the whole job:</para>
        /// <list type="bullet">
        /// <item><description><c>{"d":{"GetSiteScripts":{…}}}</c> — a method-name wrapper, so descend</description></item>
        /// <item><description><c>{"d":{"__metadata":{…},"Value":"…"}}</c> — the object itself, so stop</description></item>
        /// </list>
        /// <para>They are told apart by member count: the direct form always carries
        /// <c>__metadata</c> alongside the real members, so it never has exactly one. Descending
        /// blindly returned the <c>__metadata</c> object, and a storage entity read as null while
        /// the write that produced it had plainly succeeded.</para>
        /// </remarks>
        internal static JsonElement Unwrap(JsonElement root)
        {
            JsonElement current = root;

            if (current.ValueKind == JsonValueKind.Object
                && current.TryGetProperty("d", out JsonElement wrapper))
            {
                current = wrapper;

                if (SingleMember(current, out JsonProperty only)
                    && !string.Equals(only.Name, "__metadata", StringComparison.Ordinal))
                {
                    current = only.Value;
                }
            }

            if (current.ValueKind == JsonValueKind.Object
                && current.TryGetProperty("results", out JsonElement results))
            {
                current = results;
            }

            return current;
        }

        private static bool SingleMember(JsonElement element, out JsonProperty only)
        {
            only = default;

            if (element.ValueKind != JsonValueKind.Object)
            {
                return false;
            }

            int count = 0;

            foreach (JsonProperty property in element.EnumerateObject())
            {
                if (++count > 1)
                {
                    return false;
                }

                only = property;
            }

            return count == 1;
        }

        /// <summary>
        /// Reads a collection property, which verbose OData writes as <c>{"results":[…]}</c>.
        /// </summary>
        internal static IEnumerable<JsonElement> CollectionOf(JsonElement element, string name)
        {
            if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(name, out JsonElement property))
            {
                yield break;
            }

            JsonElement array = property.ValueKind == JsonValueKind.Object
                && property.TryGetProperty("results", out JsonElement results)
                ? results
                : property;

            if (array.ValueKind != JsonValueKind.Array)
            {
                yield break;
            }

            foreach (JsonElement item in array.EnumerateArray())
            {
                yield return item;
            }
        }

        /// <summary>
        /// Wraps a collection the way verbose OData requires on the way <em>out</em>.
        /// </summary>
        /// <remarks>
        /// A bare JSON array in a request body is rejected before it reaches the endpoint, with
        /// <c>"A node of type 'StartArray' was read from the JSON reader when trying to read a value
        /// of a property"</c> - which names no property and reads like a wrong payload rather than a
        /// wrong encoding.
        /// </remarks>
        internal static Dictionary<string, object> Collection(IEnumerable<string> values)
        {
            return new Dictionary<string, object> { ["results"] = new List<string>(values) };
        }

        internal static string StringOf(JsonElement element, string name)
        {
            return element.ValueKind == JsonValueKind.Object
                && element.TryGetProperty(name, out JsonElement value)
                && value.ValueKind == JsonValueKind.String
                ? value.GetString()
                : null;
        }

        internal static Guid GuidOf(JsonElement element, string name)
        {
            return Guid.TryParse(StringOf(element, name), out Guid id) ? id : Guid.Empty;
        }

        internal static bool BoolOf(JsonElement element, string name)
        {
            if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(name, out JsonElement value))
            {
                return false;
            }

            // Verbose OData answers with a real boolean in some places and the string "true" in
            // others - DesignType and ListColor on a site design come back as strings.
            return value.ValueKind == JsonValueKind.True
                || (value.ValueKind == JsonValueKind.String
                    && bool.TryParse(value.GetString(), out bool parsed) && parsed);
        }
    }
}
