using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using System;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.ObjectHandlers.Utilities
{
    /// <summary>
    /// Tells the modern site types apart from classic ones.
    /// </summary>
    /// <remarks>
    /// <para>Several handlers branch on this, because a setting that applies cleanly to a classic
    /// site is silently ignored - or rejected - on a modern one. PnP Framework had
    /// <c>Site.IsCommunicationSite()</c> and <c>Site.IsModernTeamSite()</c> extensions; PnP Core has
    /// neither, but both are decided by the root web's template, which PnP Core exposes.</para>
    /// <para>Deciding this from <c>WebTemplate</c> rather than from a feature probe matters: it is
    /// one already-loaded property rather than a round trip, and handlers ask this question often.</para>
    /// </remarks>
    internal static class SiteTypeHelper
    {
        /// <summary>Communication site.</summary>
        private const string CommunicationSiteTemplate = "SITEPAGEPUBLISHING";

        /// <summary>Modern, group-connected team site.</summary>
        private const string GroupTeamSiteTemplate = "GROUP";

        /// <summary>
        /// Whether the site collection is a communication site.
        /// </summary>
        internal static async Task<bool> IsCommunicationSiteAsync(PnPContext context)
        {
            return await HasRootWebTemplateAsync(context, CommunicationSiteTemplate).ConfigureAwait(false);
        }

        /// <summary>
        /// Whether the site collection is a modern, group-connected team site.
        /// </summary>
        internal static async Task<bool> IsModernTeamSiteAsync(PnPContext context)
        {
            return await HasRootWebTemplateAsync(context, GroupTeamSiteTemplate).ConfigureAwait(false);
        }

        /// <summary>
        /// Whether the site collection is a classic one - neither a communication site nor a
        /// group-connected team site.
        /// </summary>
        internal static async Task<bool> IsClassicSiteAsync(PnPContext context)
        {
            return !await IsCommunicationSiteAsync(context).ConfigureAwait(false)
                && !await IsModernTeamSiteAsync(context).ConfigureAwait(false);
        }

        private static async Task<bool> HasRootWebTemplateAsync(PnPContext context, string template)
        {
            IWeb rootWeb = await context.Site.RootWeb.GetAsync(w => w.WebTemplate).ConfigureAwait(false);

            return template.Equals(rootWeb.WebTemplate, StringComparison.InvariantCultureIgnoreCase);
        }
    }
}
