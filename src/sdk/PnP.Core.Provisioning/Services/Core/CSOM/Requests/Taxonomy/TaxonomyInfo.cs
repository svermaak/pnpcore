using System;
using System.Text.Json;

namespace PnP.Core.Provisioning.Model
{
    /// <summary>
    /// A term group, as read back from the taxonomy CSOM API.
    /// </summary>
    internal sealed class TermGroupInfo
    {
        internal Guid Id { get; set; }

        internal string Name { get; set; }

        internal string Description { get; set; }

        internal bool IsSiteCollectionGroup { get; set; }
    }

    /// <summary>
    /// A term set, as read back from the taxonomy CSOM API.
    /// </summary>
    internal sealed class TermSetInfo
    {
        internal Guid Id { get; set; }

        internal string Name { get; set; }

        internal string Description { get; set; }
    }

    /// <summary>
    /// A term, as read back from the taxonomy CSOM API.
    /// </summary>
    internal sealed class TermInfo
    {
        internal Guid Id { get; set; }

        internal string Name { get; set; }
    }
}

namespace PnP.Core.Provisioning.Services.Core.CSOM.Requests.Taxonomy
{
    /// <summary>
    /// Reads values out of a CSOM taxonomy JSON response.
    /// </summary>
    internal static class TaxonomyJson
    {
        /// <summary>
        /// Reads a GUID property, unwrapping CSOM's <c>/Guid(...)/</c> encoding.
        /// </summary>
        /// <remarks>
        /// CSOM serializes a GUID as the string <c>/Guid(3c533617-3706-4f75-8a08-8ba636f3efea)/</c>
        /// rather than as a bare GUID. Parsing it without stripping the wrapper silently yields
        /// <see cref="Guid.Empty"/> - and an empty term id looks exactly like "the server assigned
        /// its own", which is the failure S1 exists to prevent. PnP Core's own
        /// <c>GetParentOfTermRequest</c> does the same unwrapping.
        /// </remarks>
        internal static Guid ReadGuid(JsonElement element, string propertyName)
        {
            if (!element.TryGetProperty(propertyName, out JsonElement property))
            {
                return Guid.Empty;
            }

            string raw = property.GetString();
            if (string.IsNullOrEmpty(raw))
            {
                return Guid.Empty;
            }

            raw = raw.Replace("/Guid(", "").Replace(")/", "");

            return Guid.TryParse(raw, out Guid value) ? value : Guid.Empty;
        }
    }
}
