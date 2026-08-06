using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using System;
using System.Linq;
using System.Threading.Tasks;
using PnP.Core.Model;
using PnP.Core.QueryModel;

namespace PnP.Core.Provisioning.ObjectHandlers.TokenDefinitions
{
    /// <summary>
    /// Finds the term group scoped to the current site collection, which backs both
    /// <c>{sitecollectiontermgroupid}</c> and <c>{sitecollectiontermgroupname}</c>.
    /// </summary>
    /// <remarks>
    /// <para>CSOM had <c>TermStore.GetSiteCollectionGroup(site, createIfMissing)</c>, which both
    /// located the group and created it on demand. The Graph term store has no such call: groups
    /// carry a <see cref="TermGroupScope"/>, and a site-collection-scoped group has to be matched
    /// by its id, which SharePoint derives from the site id.</para>
    /// <para><b>Read-only by design.</b> The CSOM version could create the group; this one never
    /// does. Token resolution is a read operation, and a token that silently provisions taxonomy
    /// as a side effect of being parsed is a surprise nobody wants. Creating the group belongs to
    /// <c>ObjectTermGroups</c> in phase 6. Both tokens resolve to an empty string when no such
    /// group exists, matching what PnP Framework did when its lookup threw.</para>
    /// <para>Whether <see cref="TermGroupScope.SiteCollection"/> is a faithful stand-in for CSOM's
    /// <c>IsSiteCollectionGroup</c> is probe (d) of spike <b>S1</b>; the outcome may change how
    /// this resolves, but not its shape.</para>
    /// </remarks>
    internal static class SiteCollectionTermGroupResolver
    {
        internal static async Task<ITermGroup> GetAsync(PnPContext context)
        {
            try
            {
                ISite site = await context.Site.GetAsync(s => s.Id).ConfigureAwait(false);

                await context.TermStore.LoadAsync(t => t.Groups.QueryProperties(
                    g => g.Id, g => g.Name, g => g.Scope)).ConfigureAwait(false);

                // SharePoint names the site collection group's id after the site id, so match on
                // that rather than on the display name, which is localized and user-editable.
                string siteId = site.Id.ToString();

                return context.TermStore.Groups.AsRequested().FirstOrDefault(g =>
                    g.Scope == TermGroupScope.SiteCollection
                    && !string.IsNullOrEmpty(g.Id)
                    && g.Id.EndsWith(siteId, StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception)
            {
                // No term store access, or no site collection group. PnP Framework swallowed the
                // same two cases (ServerUnauthorizedAccessException and NullReferenceException)
                // and left the token unresolved.
                return null;
            }
        }
    }
}