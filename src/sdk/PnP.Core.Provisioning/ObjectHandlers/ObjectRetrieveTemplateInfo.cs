using Microsoft.Extensions.Logging;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.Services;
using System;
using System.Globalization;
using System.Text.Json;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Stamps the extracted template with its identity: id, version, scope and base site template.
    /// </summary>
    /// <remarks>
    /// <para><b>Extract only</b>, and it runs <em>after</em> <see cref="ObjectPropertyBagEntry"/>
    /// because it consumes two of the entries that handler produced. When a site was previously
    /// provisioned, <see cref="ObjectPersistTemplateInfo"/> left <c>_PnP_ProvisioningTemplateId</c>
    /// and <c>_PnP_ProvisioningTemplateInfo</c> in the property bag; this handler lifts them back
    /// into the template's own metadata and <b>removes them from the property bag entries</b>, so a
    /// re-extracted template does not carry engine bookkeeping as site data.</para>
    /// <para>Ordering therefore matters in both directions: run it before the property bag handler
    /// and there is nothing to read; skip it and every round trip accumulates two more entries.</para>
    /// </remarks>
    internal class ObjectRetrieveTemplateInfo : ObjectHandlerBase
    {
        /// <summary>
        /// Property bag key holding the id of the template last applied to this site.
        /// </summary>
        internal const string TemplateIdKey = "_PnP_ProvisioningTemplateId";

        /// <summary>
        /// Property bag key holding the serialized <see cref="ProvisioningTemplateInfo"/>.
        /// </summary>
        internal const string TemplateInfoKey = "_PnP_ProvisioningTemplateInfo";

        /// <summary>
        /// Deserialization settings for the stored template info.
        /// </summary>
        /// <remarks>
        /// PnP Framework wrote this value with Newtonsoft, which emits the property names as
        /// declared - so the names match. Case-insensitive matching is switched on anyway: the value
        /// on a real site may have been written by any version of any tool, and the cost of being
        /// tolerant here is nothing next to silently defaulting a template's version to 1.
        /// </remarks>
        private static readonly JsonSerializerOptions TemplateInfoJsonOptions = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
        };

        public ObjectRetrieveTemplateInfo()
        {
            ReportProgress = false;
        }

        public override string Name => "Retrieve Template Info";

        public override string InternalName => "RetrieveTemplateInfo";

        public override bool WillProvision(PnPContext context, ProvisioningTemplate template, ApplyConfiguration configuration)
        {
            return false;
        }

        public override bool WillExtract(PnPContext context, ProvisioningTemplate template, ExtractConfiguration configuration)
        {
            return true;
        }

        public override Task<TokenParser> ProvisionObjectsAsync(PnPContext context, ProvisioningTemplate template, TokenParser parser, ApplyConfiguration configuration)
        {
            // Nothing to apply - the write side of this pair is ObjectPersistTemplateInfo.
            return Task.FromResult(parser);
        }

        public override async Task<ProvisioningTemplate> ExtractObjectsAsync(PnPContext context, ProvisioningTemplate template, ExtractConfiguration configuration)
        {
            using (context.Logger?.BeginScope(Name))
            {
                IWeb web = context.Web;
                await web.LoadAsync(w => w.WebTemplate, w => w.WebTemplateConfiguration, w => w.ServerRelativeUrl).ConfigureAwait(false);
                await context.Site.LoadAsync(s => s.ServerRelativeUrl).ConfigureAwait(false);

                // Defaults, overridden below by anything the site itself remembers
                template.Id = $"TEMPLATE-{Guid.NewGuid():N}".ToUpper(CultureInfo.InvariantCulture);
                template.Version = 1;

                template.BaseSiteTemplate = $"{web.WebTemplate}#{BaseTemplates.BaseTemplateManager.GetConfiguration(web.WebTemplateConfiguration)}";
                template.Scope = !IsSubSite(web) ? ProvisioningTemplateScope.RootSite : ProvisioningTemplateScope.Web;

                ReadTemplateId(template);
                ReadTemplateInfo(template, context);

                return template;
            }
        }

        /// <summary>
        /// Lifts the template id out of the property bag entries and removes the entry.
        /// </summary>
        private static void ReadTemplateId(ProvisioningTemplate template)
        {
            int index = template.PropertyBagEntries.FindIndex(e => e.Key.Equals(TemplateIdKey, StringComparison.Ordinal));
            if (index < 0)
            {
                return;
            }

            string templateId = template.PropertyBagEntries[index].Value;
            if (!string.IsNullOrEmpty(templateId))
            {
                template.Id = templateId;
            }

            template.PropertyBagEntries.RemoveAt(index);
        }

        /// <summary>
        /// Lifts the serialized template info out of the property bag entries and removes the entry.
        /// </summary>
        /// <remarks>
        /// A site's property bag is not a trusted source - the value may be truncated, hand-edited
        /// or written by an older version. PnP Framework let a <c>JsonReaderException</c> escape and
        /// fail the whole extraction; here a malformed value is logged and the defaults stand, since
        /// losing the template's version number is not worth losing the template over.
        /// </remarks>
        private static void ReadTemplateInfo(ProvisioningTemplate template, PnPContext context)
        {
            int index = template.PropertyBagEntries.FindIndex(e => e.Key.Equals(TemplateInfoKey, StringComparison.Ordinal));
            if (index < 0)
            {
                return;
            }

            string json = template.PropertyBagEntries[index].Value;
            template.PropertyBagEntries.RemoveAt(index);

            if (string.IsNullOrWhiteSpace(json))
            {
                return;
            }

            ProvisioningTemplateInfo info;
            try
            {
                info = JsonSerializer.Deserialize<ProvisioningTemplateInfo>(json, TemplateInfoJsonOptions);
            }
            catch (JsonException ex)
            {
                context.Logger?.LogWarning(ex, "{Source}: {Key} could not be read and was ignored.",
                    Constants.LOGGING_SOURCE, TemplateInfoKey);
                return;
            }

            if (info == null)
            {
                return;
            }

            if (!string.IsNullOrEmpty(info.TemplateId))
            {
                template.Id = info.TemplateId;
            }

            if (!string.IsNullOrEmpty(info.TemplateSitePolicy))
            {
                template.SitePolicy = info.TemplateSitePolicy;
            }

            if (info.TemplateVersion > 0)
            {
                template.Version = info.TemplateVersion;
            }
        }
    }
}
