using PnP.Core.Services;
using System;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.ObjectHandlers.Utilities
{
    /// <summary>
    /// Tenant storage entities - the key/value pairs SharePoint stores on the app catalog site.
    /// </summary>
    /// <remarks>
    /// <para>Backlog <b>T8</b>. PnP Core has nothing for these; the REST endpoints are simple enough
    /// that CSOM would be the harder route.</para>
    /// <para><b>They live on the app catalog site, not the admin site</b>, which is the detail that
    /// makes this look broken when it is merely pointed at the wrong web: every call below succeeds
    /// against any web and silently writes a property nobody reads. The caller supplies a context
    /// already bound to the catalog, so the mistake cannot be made here.</para>
    /// </remarks>
    internal static class StorageEntities
    {
        /// <summary>
        /// Writes one storage entity.
        /// </summary>
        internal static async Task SetAsync(PnPContext appCatalog, string key, string value,
            string description, string comment)
        {
            string url = "_api/web/SetStorageEntity(" +
                $"key='{Escape(key)}'," +
                $"value='{Escape(value)}'," +
                $"description='{Escape(description)}'," +
                $"comments='{Escape(comment)}')";

            await appCatalog.Web.ExecuteRequestAsync(
                new ApiRequest(HttpMethod.Post, ApiRequestType.SPORest, url, "{}")).ConfigureAwait(false);
        }

        /// <summary>
        /// Reads one storage entity's value, or null when it is not set.
        /// </summary>
        /// <remarks>
        /// <para>An unset key is not an error - the endpoint answers with a null <c>Value</c> rather
        /// than a 404, so "missing" and "set to nothing" are the same answer and both mean null
        /// here.</para>
        /// <para>The response is unwrapped through <see cref="VerboseOData"/>: it arrives as
        /// <c>{"d":{"GetStorageEntity":{"Value":…}}}</c>. Reading <c>Value</c> off the root - which
        /// this did at first - returns null for an entity that was written successfully a moment
        /// earlier.</para>
        /// </remarks>
        internal static async Task<string> GetAsync(PnPContext appCatalog, string key)
        {
            ApiRequestResponse response = await appCatalog.Web.ExecuteRequestAsync(
                new ApiRequest(ApiRequestType.SPORest, $"_api/web/GetStorageEntity('{Escape(key)}')"))
                .ConfigureAwait(false);

            if (string.IsNullOrEmpty(response.Response))
            {
                return null;
            }

            using (JsonDocument document = JsonDocument.Parse(response.Response))
            {
                return VerboseOData.StringOf(VerboseOData.Unwrap(document.RootElement), "Value");
            }
        }

        internal static async Task RemoveAsync(PnPContext appCatalog, string key)
        {
            await appCatalog.Web.ExecuteRequestAsync(new ApiRequest(HttpMethod.Post, ApiRequestType.SPORest,
                $"_api/web/RemoveStorageEntity(key='{Escape(key)}')", "{}")).ConfigureAwait(false);
        }

        private static string Escape(string value)
        {
            return (value ?? string.Empty).Replace("'", "''");
        }
    }
}
