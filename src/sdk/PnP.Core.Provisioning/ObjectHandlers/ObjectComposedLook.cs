using Microsoft.Extensions.Logging;
using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.Provisioning.ObjectHandlers.Utilities;
using System.Net.Http;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using ComposedLookModel = PnP.Core.Provisioning.Model.ComposedLook;
using CoreList = PnP.Core.Model.SharePoint.IList;

namespace PnP.Core.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Applies and reads back the <c>&lt;pnp:ComposedLook&gt;</c> element - a classic site's colour
    /// palette, font scheme, background image and master page.
    /// </summary>
    /// <remarks>
    /// <para>Backlog <b>T19</b>. This is the <b>classic</b> theming mechanism, and PnP Core has no
    /// model for it: <c>IBrandingManager</c> covers modern themes only. A composed look is an item
    /// in the site's Design Catalog pointing at three files, applied through <c>ApplyTheme</c>.</para>
    /// <para><b>No CSOM request was needed</b>, which is what T19's audit predicted:
    /// <c>ApplyTheme</c> is exposed over REST, the design catalog is an ordinary list, and
    /// <c>MasterUrl</c> / <c>CustomMasterUrl</c> are settable on <c>IWeb</c>.</para>
    /// <para><b>Composed looks do not exist on a NoScript site</b>, which is every modern site. The
    /// handler declines with a warning there; that is not a gap, it is the feature's scope.</para>
    /// <para><b>What was applied is recorded in the property bag.</b> SharePoint does not remember
    /// which named look produced the current theme - it stores the resolved file urls and nothing
    /// else - so extraction cannot recover the name. PnP Framework wrote the whole element into
    /// <c>_PnP_ProvisioningTemplateComposedLookInfo</c> on apply and read it back on extract, and
    /// that round trip is preserved here. Without it an extract-then-apply cycle silently loses the
    /// look's identity.</para>
    /// </remarks>
    internal class ObjectComposedLook : ObjectHandlerBase
    {
        /// <summary>
        /// The property bag entry PnP Framework records the applied look in.
        /// </summary>
        /// <remarks>
        /// The key is part of the interchange format, not an implementation detail: a site
        /// provisioned by PnP Framework and extracted by this engine has to find it, and vice versa.
        /// </remarks>
        private const string ComposedLookInfoKey = "_PnP_ProvisioningTemplateComposedLookInfo";

        public override string Name => "Composed look";

        public override string InternalName => "ComposedLooks";

        public override bool WillProvision(PnPContext context, ProvisioningTemplate template, ApplyConfiguration configuration)
        {
            _willProvision ??= template.ComposedLook != null && !IsEmpty(template.ComposedLook);
            return _willProvision.Value;
        }

        public override bool WillExtract(PnPContext context, ProvisioningTemplate template, ExtractConfiguration configuration)
        {
            _willExtract ??= true;
            return _willExtract.Value;
        }

        /// <summary>
        /// Whether the element says nothing worth applying.
        /// </summary>
        /// <remarks>
        /// An extracted template from a site with no composed look carries the element with
        /// <c>Name="Current"</c> and no files, which means "whatever the site already has".
        /// Re-applying that is a no-op at best.
        /// </remarks>
        private static bool IsEmpty(ComposedLookModel look)
        {
            return string.IsNullOrEmpty(look.ColorFile)
                && string.IsNullOrEmpty(look.FontFile)
                && string.IsNullOrEmpty(look.BackgroundFile)
                && (string.IsNullOrEmpty(look.Name)
                    || look.Name.Equals("Current", StringComparison.OrdinalIgnoreCase));
        }

        #region Apply

        public override async Task<TokenParser> ProvisionObjectsAsync(PnPContext context, ProvisioningTemplate template,
            TokenParser parser, ApplyConfiguration configuration)
        {
            ComposedLookModel look = template.ComposedLook;

            if (look == null || IsEmpty(look))
            {
                return parser;
            }

            if (await context.Web.IsNoScriptSiteAsync().ConfigureAwait(false))
            {
                string warning = "This is a NoScript site, so classic composed looks are not supported. " +
                    "The composed look was skipped; use <pnp:Theme> for a modern theme instead.";
                context.Logger?.LogWarning("{Source}: {Message}", Constants.LOGGING_SOURCE, warning);
                WriteMessage(warning, ProvisioningMessageType.Warning);
                return parser;
            }

            string colorFile = parser.ParseString(look.ColorFile);
            string fontFile = parser.ParseString(look.FontFile);
            string backgroundFile = parser.ParseString(look.BackgroundFile);
            string masterUrl = parser.ParseString(template.WebSettings?.MasterPageUrl);

            try
            {
                if (string.IsNullOrEmpty(colorFile) && string.IsNullOrEmpty(fontFile) && string.IsNullOrEmpty(backgroundFile))
                {
                    // A name with no files means an existing entry in the Design Catalog, so the
                    // files come from that entry rather than from the template.
                    (colorFile, fontFile, backgroundFile, string catalogMaster) =
                        await ReadFromCatalogAsync(context, parser.ParseString(look.Name)).ConfigureAwait(false);

                    if (string.IsNullOrEmpty(masterUrl))
                    {
                        masterUrl = catalogMaster;
                    }

                    if (string.IsNullOrEmpty(colorFile))
                    {
                        string warning = $"The composed look '{look.Name}' is not in this site's design catalog, " +
                            "and the template supplies no files of its own, so nothing was applied.";
                        context.Logger?.LogWarning("{Source}: {Message}", Constants.LOGGING_SOURCE, warning);
                        WriteMessage(warning, ProvisioningMessageType.Warning);
                        return parser;
                    }
                }

                await ApplyMasterPageAsync(context, masterUrl).ConfigureAwait(false);
                await ApplyThemeAsync(context, colorFile, fontFile, backgroundFile).ConfigureAwait(false);
                await RecordAppliedLookAsync(context, look).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                string warning = $"The composed look '{look.Name}' could not be applied: {ErrorText.Describe(ex)}";
                context.Logger?.LogWarning(ex, "{Source}: {Message}", Constants.LOGGING_SOURCE, warning);
                WriteMessage(warning, ProvisioningMessageType.Warning);
            }

            return parser;
        }

        /// <summary>
        /// Applies the palette, font scheme and background image.
        /// </summary>
        /// <remarks>
        /// <para><b>REST, not CSOM.</b> <c>ApplyTheme</c> is one of the few classic operations that
        /// SharePoint does expose over REST, which is what the <b>T19</b> audit concluded and what
        /// makes this handler need no new CSOM request at all - only the property bag round trip
        /// below is bespoke.</para>
        /// <para>Each url is a method parameter in the endpoint's path. An unspecified one is the
        /// literal <c>null</c>, unquoted: an empty string is a url, and SharePoint goes looking for
        /// a file at the site root rather than treating the scheme as absent.</para>
        /// <para><c>sharegenerated=true</c> puts the generated theme files in the site collection's
        /// shared location rather than this web's, which is what lets a subsite inherit them.</para>
        /// </remarks>
        private static async Task ApplyThemeAsync(PnPContext context, string colorFile, string fontFile, string backgroundFile)
        {
            string request = "_api/web/applytheme"
                + $"(colorpaletteurl={Quote(colorFile)}"
                + $",fontschemeurl={Quote(fontFile)}"
                + $",backgroundimageurl={Quote(backgroundFile)}"
                + ",sharegenerated=true)";

            await context.Web.ExecuteRequestAsync(new ApiRequest(HttpMethod.Post, ApiRequestType.SPORest, request, null))
                .ConfigureAwait(false);
        }

        private static string Quote(string url)
        {
            return string.IsNullOrWhiteSpace(url)
                ? "null"
                : "'" + url.Replace("'", "''") + "'";
        }

        private static async Task ApplyMasterPageAsync(PnPContext context, string masterUrl)
        {
            if (string.IsNullOrEmpty(masterUrl))
            {
                return;
            }

            IWeb web = context.Web;
            await web.LoadAsync(w => w.MasterUrl, w => w.CustomMasterUrl).ConfigureAwait(false);

            bool dirty = false;

            if (!string.Equals(web.MasterUrl, masterUrl, StringComparison.OrdinalIgnoreCase))
            {
                web.MasterUrl = masterUrl;
                dirty = true;
            }

            if (!string.Equals(web.CustomMasterUrl, masterUrl, StringComparison.OrdinalIgnoreCase))
            {
                web.CustomMasterUrl = masterUrl;
                dirty = true;
            }

            if (dirty)
            {
                await web.UpdateAsync().ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Records the applied look so a later extract can recover its name.
        /// </summary>
        private async Task RecordAppliedLookAsync(PnPContext context, ComposedLookModel look)
        {
            try
            {
                IWeb web = context.Web;
                await web.LoadAsync(w => w.AllProperties).ConfigureAwait(false);

                web.AllProperties[ComposedLookInfoKey] = JsonSerializer.Serialize(new StoredComposedLook
                {
                    Name = look.Name,
                    ColorFile = look.ColorFile,
                    FontFile = look.FontFile,
                    BackgroundFile = look.BackgroundFile,
                    Version = look.Version,
                });

                await web.AllProperties.UpdateAsync().ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                // Losing the record costs the look's name on a later extract, not the look itself.
                context.Logger?.LogDebug(ex, "{Source}: the applied composed look could not be recorded.",
                    Constants.LOGGING_SOURCE);
            }
        }

        #endregion

        #region Extract

        public override async Task<ProvisioningTemplate> ExtractObjectsAsync(PnPContext context, ProvisioningTemplate template,
            ExtractConfiguration configuration)
        {
            ComposedLookModel recorded = await ReadRecordedLookAsync(context).ConfigureAwait(false);

            if (recorded != null)
            {
                template.ComposedLook = recorded;
                return template;
            }

            // Nothing recorded. SharePoint keeps the resolved file urls but not the look's name, so
            // the best that can be reported is "whatever this site currently has" - which is what
            // PnP Framework emitted too.
            template.ComposedLook = new ComposedLookModel { Name = "Current" };

            return template;
        }

        private async Task<ComposedLookModel> ReadRecordedLookAsync(PnPContext context)
        {
            try
            {
                IWeb web = context.Web;
                await web.LoadAsync(w => w.AllProperties, w => w.Url).ConfigureAwait(false);

                if (!web.AllProperties.Values.TryGetValue(ComposedLookInfoKey, out object raw)
                    || raw == null
                    || string.IsNullOrWhiteSpace(raw.ToString()))
                {
                    return null;
                }

                StoredComposedLook stored = JsonSerializer.Deserialize<StoredComposedLook>(raw.ToString());

                if (stored == null || string.IsNullOrEmpty(stored.Name))
                {
                    return null;
                }

                string webUrl = web.Url.ToString();

                return new ComposedLookModel
                {
                    Name = stored.Name,
                    ColorFile = Tokenize(stored.ColorFile, webUrl),
                    FontFile = Tokenize(stored.FontFile, webUrl),
                    BackgroundFile = Tokenize(stored.BackgroundFile, webUrl),
                    Version = stored.Version,
                };
            }
            catch (Exception ex)
            {
                context.Logger?.LogDebug(ex, "{Source}: the recorded composed look could not be read.",
                    Constants.LOGGING_SOURCE);
                return null;
            }
        }

        #endregion

        #region Design catalog

        /// <summary>
        /// Reads the file urls of a named entry in the site's Design Catalog.
        /// </summary>
        /// <remarks>
        /// The catalog is list template 124. Its <c>ThemeUrl</c>, <c>FontSchemeUrl</c>,
        /// <c>ImageUrl</c> and <c>MasterPageUrl</c> columns are url fields holding absolute urls,
        /// and <c>ApplyTheme</c> wants server relative ones.
        /// </remarks>
        private static async Task<(string Color, string Font, string Background, string Master)> ReadFromCatalogAsync(
            PnPContext context, string lookName)
        {
            if (string.IsNullOrEmpty(lookName))
            {
                return (null, null, null, null);
            }

            await context.Web.LoadAsync(w => w.Lists.QueryProperties(l => l.Id, l => l.Title, l => l.TemplateType))
                .ConfigureAwait(false);

            CoreList catalog = context.Web.Lists.AsRequested()
                .FirstOrDefault(l => l.TemplateType == ListTemplateType.DesignCatalog);

            if (catalog == null)
            {
                return (null, null, null, null);
            }

            string query = "<View><Query><Where><Eq><FieldRef Name=\"Name\"/>" +
                $"<Value Type=\"Text\">{System.Security.SecurityElement.Escape(lookName)}</Value>" +
                "</Eq></Where></Query><RowLimit>1</RowLimit></View>";

            await catalog.LoadItemsByCamlQueryAsync(query).ConfigureAwait(false);

            IListItem item = catalog.Items.AsRequested().FirstOrDefault();

            if (item == null)
            {
                return (null, null, null, null);
            }

            return (
                ServerRelative(item, "ThemeUrl"),
                ServerRelative(item, "FontSchemeUrl"),
                ServerRelative(item, "ImageUrl"),
                ServerRelative(item, "MasterPageUrl"));
        }

        private static string ServerRelative(IListItem item, string fieldName)
        {
            if (!item.Values.TryGetValue(fieldName, out object value) || value == null)
            {
                return null;
            }

            string url = value is IFieldUrlValue urlValue ? urlValue.Url : value.ToString();

            if (string.IsNullOrEmpty(url))
            {
                return null;
            }

            // The catalog stores absolute urls; ApplyTheme takes server relative ones, and the
            // difference is not something SharePoint corrects for.
            return Uri.TryCreate(url, UriKind.Absolute, out Uri absolute)
                ? Uri.UnescapeDataString(absolute.AbsolutePath)
                : Uri.UnescapeDataString(url);
        }

        #endregion

        /// <summary>
        /// The shape written into the property bag.
        /// </summary>
        /// <remarks>
        /// Deliberately its own type rather than serializing <see cref="ComposedLookModel"/>: the
        /// model carries base-class members that would end up in the JSON, and the entry has to stay
        /// readable by PnP Framework, which wrote exactly these five properties.
        /// </remarks>
        private sealed class StoredComposedLook
        {
            public string Name { get; set; }

            public string ColorFile { get; set; }

            public string FontFile { get; set; }

            public string BackgroundFile { get; set; }

            public int Version { get; set; }
        }
    }
}
