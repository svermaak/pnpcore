using Microsoft.Extensions.Logging;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.Services;
using System;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Provisions and extracts the search configuration XML at site and web scope.
    /// </summary>
    /// <remarks>
    /// One of the cheapest handlers to port: PnP Core already implements search configuration
    /// import and export through CSOM internally, so this is a straight call on each side rather
    /// than anything the migration had to build.
    /// </remarks>
    internal class ObjectSearchSettings : ObjectHandlerBase
    {
        public override string Name => "Search Settings";

        public override string InternalName => "SearchSettings";

        public override bool WillExtract(PnPContext context, ProvisioningTemplate template, ExtractConfiguration configuration)
        {
            return true;
        }

        public override bool WillProvision(PnPContext context, ProvisioningTemplate template, ApplyConfiguration configuration)
        {
            return !string.IsNullOrEmpty(template.SiteSearchSettings)
                || !string.IsNullOrEmpty(template.WebSearchSettings);
        }

        public override async Task<ProvisioningTemplate> ExtractObjectsAsync(PnPContext context, ProvisioningTemplate template, ExtractConfiguration configuration)
        {
            using (context.Logger?.BeginScope(Name))
            {
                // Search configuration is unavailable on some site types and under some permission
                // sets. PnP Framework swallowed ServerException here; the same tolerance applies -
                // a site with no search customization should extract cleanly, not fail.
                try
                {
                    string siteSearchSettings = await context.Site.GetSearchConfigurationXmlAsync().ConfigureAwait(false);
                    if (!string.IsNullOrEmpty(siteSearchSettings))
                    {
                        template.SiteSearchSettings = siteSearchSettings;
                    }
                }
                catch (Exception ex)
                {
                    context.Logger?.LogWarning(ex, "{Source}: could not read the site search configuration", Constants.LOGGING_SOURCE);
                }

                try
                {
                    string webSearchSettings = await context.Web.GetSearchConfigurationXmlAsync().ConfigureAwait(false);
                    if (!string.IsNullOrEmpty(webSearchSettings))
                    {
                        template.WebSearchSettings = webSearchSettings;
                    }
                }
                catch (Exception ex)
                {
                    context.Logger?.LogWarning(ex, "{Source}: could not read the web search configuration", Constants.LOGGING_SOURCE);
                }

                return template;
            }
        }

        public override async Task<TokenParser> ProvisionObjectsAsync(PnPContext context, ProvisioningTemplate template, TokenParser parser, ApplyConfiguration configuration)
        {
            using (context.Logger?.BeginScope(Name))
            {
                if (!string.IsNullOrEmpty(template.SiteSearchSettings))
                {
                    await context.Site.SetSearchConfigurationXmlAsync(
                        parser.ParseString(template.SiteSearchSettings)).ConfigureAwait(false);
                }

                if (!string.IsNullOrEmpty(template.WebSearchSettings))
                {
                    await context.Web.SetSearchConfigurationXmlAsync(
                        parser.ParseString(template.WebSearchSettings)).ConfigureAwait(false);
                }

                return parser;
            }
        }
    }
}
