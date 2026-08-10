using Microsoft.Extensions.Logging;
using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.Provisioning.ObjectHandlers.Utilities;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using CoreList = PnP.Core.Model.SharePoint.IList;
using PageLayoutModel = PnP.Core.Provisioning.Model.PageLayout;
using PublishingModel = PnP.Core.Provisioning.Model.Publishing;

namespace PnP.Core.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Applies and reads back the <c>&lt;pnp:Publishing&gt;</c> element - which web templates a
    /// subsite may use, and which page layouts a publishing site offers.
    /// </summary>
    /// <remarks>
    /// <para>Backlog <b>T17</b>. <b>None of this is an API.</b> SharePoint stores all three settings
    /// as XML in the web's property bag - <c>__WebTemplates</c>, <c>__PageLayouts</c> and
    /// <c>__DefaultPageLayout</c> - and the "extension methods" PnP Framework calls are just
    /// builders for that XML. So the port is not a CSOM question at all; it is a matter of writing
    /// the same XML, byte for byte, because the SharePoint UI reads it back.</para>
    /// <para><b>A page layout is identified by its unique id, not its path.</b> The
    /// <c>&lt;layout&gt;</c> element carries the file's <c>UniqueId</c> and its site-relative url,
    /// and both come from the item in the root web's master page gallery - which is why layouts are
    /// resolved against the <em>root</em> web even when applying to a subsite.</para>
    /// <para><b>Design packages are not installed.</b> <c>Site.InstallSolution</c> is a sandboxed
    /// solution deployment, which is deprecated in SharePoint Online. The element is reported rather
    /// than silently ignored.</para>
    /// </remarks>
    internal class ObjectPublishing : ObjectHandlerBase
    {
        private const string AvailableWebTemplatesKey = PublishingPropertyBagXml.AvailableWebTemplatesKey;
        private const string InheritWebTemplatesKey = PublishingPropertyBagXml.InheritWebTemplatesKey;
        private const string AvailablePageLayoutsKey = PublishingPropertyBagXml.AvailablePageLayoutsKey;
        private const string DefaultPageLayoutKey = PublishingPropertyBagXml.DefaultPageLayoutKey;

        /// <summary>The master page gallery, where page layouts live.</summary>
        private const int MasterPageGalleryTemplateType = 116;

        public override string Name => "Publishing";

        public override string InternalName => "Publishing";

        public override bool WillProvision(PnPContext context, ProvisioningTemplate template, ApplyConfiguration configuration)
        {
            _willProvision ??= template.Publishing != null;
            return _willProvision.Value;
        }

        public override bool WillExtract(PnPContext context, ProvisioningTemplate template, ExtractConfiguration configuration)
        {
            _willExtract ??= true;
            return _willExtract.Value;
        }

        #region Apply

        public override async Task<TokenParser> ProvisionObjectsAsync(PnPContext context, ProvisioningTemplate template,
            TokenParser parser, ApplyConfiguration configuration)
        {
            PublishingModel publishing = template.Publishing;

            if (publishing == null)
            {
                return parser;
            }

            if (await context.Web.IsNoScriptSiteAsync().ConfigureAwait(false))
            {
                string warning = "This is a NoScript site, so the publishing settings were skipped.";
                context.Logger?.LogWarning("{Source}: {Message}", Constants.LOGGING_SOURCE, warning);
                WriteMessage(warning, ProvisioningMessageType.Warning);
                return parser;
            }

            if (!await PublishingRequirements.EnsureAsync(context, publishing.AutoCheckRequirements,
                m => WriteMessage(m, ProvisioningMessageType.Warning), "publishing settings").ConfigureAwait(false))
            {
                return parser;
            }

            if (publishing.DesignPackage != null)
            {
                string warning = "The template carries a design package. Sandboxed solution deployment is deprecated " +
                    "in SharePoint Online and this engine does not install one, so it was skipped.";
                context.Logger?.LogWarning("{Source}: {Message}", Constants.LOGGING_SOURCE, warning);
                WriteMessage(warning, ProvisioningMessageType.Warning);
            }

            await ApplyAvailableWebTemplatesAsync(context, publishing, parser).ConfigureAwait(false);
            await ApplyPageLayoutsAsync(context, publishing, parser).ConfigureAwait(false);

            return parser;
        }

        /// <summary>
        /// Writes the <c>__WebTemplates</c> property bag XML.
        /// </summary>
        /// <remarks>
        /// The XML shape lives in <see cref="PublishingPropertyBagXml"/>, which offline tests cover;
        /// what is here is the property-bag write. <c>__InheritWebTemplates</c> has to be turned off
        /// alongside, or the list is ignored in favour of the parent's.
        /// </remarks>
        private async Task ApplyAvailableWebTemplatesAsync(PnPContext context, PublishingModel publishing, TokenParser parser)
        {
            if (!publishing.AvailableWebTemplates.Any())
            {
                return;
            }

            try
            {
                await SetPropertiesAsync(context, new Dictionary<string, string>
                {
                    [AvailableWebTemplatesKey] = PublishingPropertyBagXml.BuildWebTemplates(
                        publishing.AvailableWebTemplates, parser.ParseString),

                    // Without this the list is ignored in favour of the parent's.
                    [InheritWebTemplatesKey] = "False",
                }).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                string warning = $"The available web templates could not be set: {ErrorText.Describe(ex)}";
                context.Logger?.LogWarning(ex, "{Source}: {Message}", Constants.LOGGING_SOURCE, warning);
                WriteMessage(warning, ProvisioningMessageType.Warning);
            }
        }

        private async Task ApplyPageLayoutsAsync(PnPContext context, PublishingModel publishing, TokenParser parser)
        {
            if (!publishing.PageLayouts.Any())
            {
                return;
            }

            Dictionary<string, PageLayoutEntry> catalog = await ReadPageLayoutCatalogAsync(context).ConfigureAwait(false);

            if (catalog.Count == 0)
            {
                string warning = "The master page gallery could not be read, so the page layouts were not set.";
                context.Logger?.LogWarning("{Source}: {Message}", Constants.LOGGING_SOURCE, warning);
                WriteMessage(warning, ProvisioningMessageType.Warning);
                return;
            }

            var available = new List<XElement>();
            XElement defaultLayout = null;

            foreach (PageLayoutModel layout in publishing.PageLayouts)
            {
                string path = parser.ParseString(layout.Path);

                if (!catalog.TryGetValue(NameOf(path), out PageLayoutEntry entry))
                {
                    string warning = $"The page layout '{path}' is not in the master page gallery, so it was skipped.";
                    context.Logger?.LogWarning("{Source}: {Message}", Constants.LOGGING_SOURCE, warning);
                    WriteMessage(warning, ProvisioningMessageType.Warning);
                    continue;
                }

                XElement element = PublishingPropertyBagXml.BuildLayout(entry.UniqueId, entry.SiteRelativeUrl);
                available.Add(element);

                if (layout.IsDefault)
                {
                    defaultLayout = element;
                }
            }

            if (available.Count == 0)
            {
                return;
            }

            var properties = new Dictionary<string, string>
            {
                [AvailablePageLayoutsKey] = PublishingPropertyBagXml.BuildPageLayouts(available),
            };

            if (defaultLayout != null)
            {
                properties[DefaultPageLayoutKey] = defaultLayout.ToString(SaveOptions.DisableFormatting);
            }

            try
            {
                await SetPropertiesAsync(context, properties).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                string warning = $"The available page layouts could not be set: {ErrorText.Describe(ex)}";
                context.Logger?.LogWarning(ex, "{Source}: {Message}", Constants.LOGGING_SOURCE, warning);
                WriteMessage(warning, ProvisioningMessageType.Warning);
            }
        }

        private static async Task SetPropertiesAsync(PnPContext context, Dictionary<string, string> properties)
        {
            IWeb web = context.Web;
            await web.LoadAsync(w => w.AllProperties).ConfigureAwait(false);

            foreach (KeyValuePair<string, string> property in properties)
            {
                web.AllProperties[property.Key] = property.Value;
            }

            await web.AllProperties.UpdateAsync().ConfigureAwait(false);
        }

        #endregion

        #region Extract

        public override async Task<ProvisioningTemplate> ExtractObjectsAsync(PnPContext context, ProvisioningTemplate template,
            ExtractConfiguration configuration)
        {
            if (!await PublishingRequirements.IsWebPublishingActiveAsync(context).ConfigureAwait(false))
            {
                // Not a publishing site. Emitting an empty element would make every modern site's
                // template carry one for no reason.
                return template;
            }

            // ObjectImageRenditions writes into the same element, and either may run first.
            template.Publishing ??= new PublishingModel();
            template.Publishing.AutoCheckRequirements = AutoCheckRequirementsOptions.MakeCompliant;

            await ExtractAvailableWebTemplatesAsync(context, template.Publishing).ConfigureAwait(false);
            await ExtractPageLayoutsAsync(context, template.Publishing).ConfigureAwait(false);

            return template;
        }

        private static async Task ExtractAvailableWebTemplatesAsync(PnPContext context, PublishingModel publishing)
        {
            try
            {
                IWeb web = context.Web;
                await web.LoadAsync(w => w.AllProperties).ConfigureAwait(false);

                if (!web.AllProperties.Values.TryGetValue(AvailableWebTemplatesKey, out object raw))
                {
                    return;
                }

                foreach (AvailableWebTemplate webTemplate in PublishingPropertyBagXml.ReadWebTemplates(raw?.ToString()))
                {
                    publishing.AvailableWebTemplates.Add(webTemplate);
                }
            }
            catch (Exception ex)
            {
                context.Logger?.LogDebug(ex, "{Source}: the available web templates could not be read.",
                    Constants.LOGGING_SOURCE);
            }
        }

        private static async Task ExtractPageLayoutsAsync(PnPContext context, PublishingModel publishing)
        {
            try
            {
                IWeb web = context.Web;
                await web.LoadAsync(w => w.AllProperties).ConfigureAwait(false);

                web.AllProperties.Values.TryGetValue(DefaultPageLayoutKey, out object rawDefault);

                if (!web.AllProperties.Values.TryGetValue(AvailablePageLayoutsKey, out object raw))
                {
                    return;
                }

                foreach (PageLayoutModel layout in PublishingPropertyBagXml.ReadPageLayouts(
                    raw?.ToString(), rawDefault?.ToString()))
                {
                    publishing.PageLayouts.Add(layout);
                }
            }
            catch (Exception ex)
            {
                context.Logger?.LogDebug(ex, "{Source}: the available page layouts could not be read.",
                    Constants.LOGGING_SOURCE);
            }
        }

        #endregion

        #region Page layout catalog

        /// <summary>
        /// One page layout in the root web's master page gallery.
        /// </summary>
        private sealed class PageLayoutEntry
        {
            internal string UniqueId { get; set; }

            internal string SiteRelativeUrl { get; set; }
        }

        /// <summary>
        /// Reads the page layouts from the <b>root</b> web's master page gallery, keyed by file name.
        /// </summary>
        /// <remarks>
        /// The root web, even when applying to a subsite: page layouts live in the site collection's
        /// gallery and the property bag XML addresses them by an id from there. Keyed by file name
        /// because a template's <c>Path</c> may be a bare name or a full url.
        /// </remarks>
        private static async Task<Dictionary<string, PageLayoutEntry>> ReadPageLayoutCatalogAsync(PnPContext context)
        {
            var catalog = new Dictionary<string, PageLayoutEntry>(StringComparer.OrdinalIgnoreCase);

            try
            {
                // Site.RootWeb, not a cloned context: PnP Core reaches the root web through the same
                // context, and the rest of this project already does it that way.
                IWeb rootWeb = await context.Site.RootWeb.GetAsync(
                    w => w.ServerRelativeUrl,
                    w => w.Lists.QueryProperties(l => l.Id, l => l.TemplateType)).ConfigureAwait(false);

                CoreList gallery = rootWeb.Lists.AsRequested()
                    .FirstOrDefault(l => (int)l.TemplateType == MasterPageGalleryTemplateType);

                if (gallery == null)
                {
                    return catalog;
                }

                string webUrl = rootWeb.ServerRelativeUrl.TrimEnd('/');

                await gallery.LoadItemsByCamlQueryAsync(
                    "<View Scope=\"RecursiveAll\"><Query></Query></View>").ConfigureAwait(false);

                foreach (IListItem item in gallery.Items.AsRequested())
                {
                    string fileRef = ValueOf(item, "FileRef");
                    string uniqueId = ValueOf(item, "UniqueId");

                    if (string.IsNullOrEmpty(fileRef) || string.IsNullOrEmpty(uniqueId))
                    {
                        continue;
                    }

                    catalog[NameOf(fileRef)] = new PageLayoutEntry
                    {
                        UniqueId = uniqueId.Trim('{', '}'),

                        // Site relative, as the property bag XML expects - the gallery stores a
                        // server relative FileRef and SharePoint does not reconcile the two.
                        SiteRelativeUrl = fileRef.StartsWith(webUrl, StringComparison.OrdinalIgnoreCase)
                            ? fileRef.Substring(webUrl.Length).TrimStart('/')
                            : fileRef.TrimStart('/'),
                    };
                }
            }
            catch (Exception ex)
            {
                context.Logger?.LogDebug(ex, "{Source}: the master page gallery could not be read.",
                    Constants.LOGGING_SOURCE);
            }

            return catalog;
        }

        private static string ValueOf(IListItem item, string fieldName)
        {
            return item.Values.TryGetValue(fieldName, out object value) ? value?.ToString() : null;
        }

        private static string NameOf(string pathOrName)
        {
            if (string.IsNullOrEmpty(pathOrName))
            {
                return string.Empty;
            }

            int lastSlash = pathOrName.Replace('\\', '/').LastIndexOf('/');
            return lastSlash < 0 ? pathOrName : pathOrName.Substring(lastSlash + 1);
        }

        #endregion
    }
}
