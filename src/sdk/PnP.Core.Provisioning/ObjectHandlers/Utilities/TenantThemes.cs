using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.ObjectHandlers.Utilities
{
    /// <summary>
    /// Tenant theme CRUD, over the <c>thememanager</c> REST endpoints.
    /// </summary>
    /// <remarks>
    /// <para>Part of backlog <b>T13</b>. PnP Core already reads these -
    /// <c>IBrandingManager.GetAvailableThemesAsync()</c> calls
    /// <c>_api/thememanager/GetTenantThemingOptions</c> - so only the write half is missing, and it
    /// is on the same controller.</para>
    /// <para><b>The palette is passed through, not interpreted.</b> A template carries the palette as
    /// a JSON string; parsing it into a dictionary only to serialise it again would turn an invalid
    /// palette into a deserialisation error in this engine rather than a validation error from
    /// SharePoint, which is the one that can say what is wrong with it. It is still parsed far enough
    /// to fail early on JSON that is not an object.</para>
    /// </remarks>
    internal static class TenantThemes
    {
        /// <summary>
        /// The theme names the tenant already has.
        /// </summary>
        internal static async Task<HashSet<string>> GetNamesAsync(PnPContext context)
        {
            List<ITheme> themes = await context.Web.GetBrandingManager()
                .GetAvailableThemesAsync().ConfigureAwait(false);

            return new HashSet<string>(
                themes.Where(t => t.Name != null).Select(t => t.Name),
                StringComparer.OrdinalIgnoreCase);
        }

        internal static async Task AddAsync(PnPContext context, string name, string palette, bool isInverted)
        {
            await PostAsync(context, "AddTenantTheme", name, palette, isInverted).ConfigureAwait(false);
        }

        internal static async Task UpdateAsync(PnPContext context, string name, string palette, bool isInverted)
        {
            await PostAsync(context, "UpdateTenantTheme", name, palette, isInverted).ConfigureAwait(false);
        }

        internal static async Task DeleteAsync(PnPContext context, string name)
        {
            string body = JsonSerializer.Serialize(new Dictionary<string, object> { ["name"] = name });

            await context.Web.ExecuteRequestAsync(new ApiRequest(HttpMethod.Post, ApiRequestType.SPORest,
                "_api/thememanager/DeleteTenantTheme", body)).ConfigureAwait(false);
        }

        private static async Task PostAsync(PnPContext context, string method, string name,
            string palette, bool isInverted)
        {
            string body = JsonSerializer.Serialize(new Dictionary<string, object>
            {
                ["name"] = name,
                ["themeJson"] = BuildThemeJson(name, palette, isInverted),
            });

            await context.Web.ExecuteRequestAsync(new ApiRequest(HttpMethod.Post, ApiRequestType.SPORest,
                $"_api/thememanager/{method}", body)).ConfigureAwait(false);
        }

        /// <summary>
        /// Wraps a palette into the shape <c>thememanager</c> expects.
        /// </summary>
        /// <remarks>
        /// <c>{ "name": ..., "palette": { ... }, "isInverted": ... }</c> - the same three members
        /// PnP Framework's <c>TenantTheme</c> contract has, and the property names are lower camel
        /// case because SharePoint matches on them exactly.
        /// </remarks>
        internal static string BuildThemeJson(string name, string palette, bool isInverted)
        {
            using (JsonDocument parsed = JsonDocument.Parse(string.IsNullOrWhiteSpace(palette) ? "{}" : palette))
            {
                if (parsed.RootElement.ValueKind != JsonValueKind.Object)
                {
                    throw new ArgumentException(
                        $"The palette for theme '{name}' is not a JSON object.", nameof(palette));
                }

                // Composed by hand rather than through a serialiser: the palette is already JSON and
                // round-tripping it through a Dictionary<string, string> would reject any palette
                // with a non-string value rather than letting SharePoint judge it.
                return "{\"name\":" + JsonSerializer.Serialize(name)
                    + ",\"palette\":" + parsed.RootElement.GetRawText()
                    + ",\"isInverted\":" + (isInverted ? "true" : "false") + "}";
            }
        }
    }
}
