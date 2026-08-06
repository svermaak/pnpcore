using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using System;
using System.Linq;
using System.Threading.Tasks;
using PnP.Core.Model;
using PnP.Core.QueryModel;

namespace PnP.Core.Provisioning.ObjectHandlers.Utilities
{
    /// <summary>
    /// Taxonomy lookups the provisioning engine needs but the Graph term store does not offer
    /// directly.
    /// </summary>
    /// <remarks>
    /// <para><b>A genuine CSOM/Graph divergence.</b> CSOM could fetch a term set by id alone -
    /// <c>TermStore.GetTermSet(Guid)</c> - because the term store was a flat, addressable object
    /// model. PnP Core's Graph term store exposes sets only underneath their group
    /// (<c>TermStore.Groups[x].Sets</c>), with no store-level set collection, so finding a set
    /// from just its id means walking the groups.</para>
    /// <para>That matters because taxonomy field XML embeds a bare <c>TermSetId</c> and nothing
    /// else - it is exactly the id-only lookup Graph does not provide. The walk is cached per
    /// context for the lifetime of a provisioning run, which keeps it to one enumeration rather
    /// than one per field.</para>
    /// <para>Recorded as a finding for spike <b>S1</b>: not a blocker, but a cost that shapes how
    /// <c>ObjectField</c> should batch its work in phase 6.</para>
    /// </remarks>
    internal static class TaxonomyLookup
    {
        /// <summary>
        /// Finds a term set and its owning group by the term set's id.
        /// </summary>
        /// <returns>The group and set, or <c>(null, null)</c> when no such set exists.</returns>
        internal static async Task<(ITermGroup Group, ITermSet Set)> FindTermSetAsync(PnPContext context, string termSetId)
        {
            if (string.IsNullOrEmpty(termSetId))
            {
                return (null, null);
            }

            await context.TermStore.LoadAsync(t => t.Groups.QueryProperties(
                g => g.Id,
                g => g.Name,
                g => g.Scope,
                g => g.Sets.QueryProperties(s => s.Id, s => s.LocalizedNames))).ConfigureAwait(false);

            foreach (ITermGroup group in context.TermStore.Groups.AsRequested())
            {
                ITermSet set = group.Sets.AsRequested()
                    .FirstOrDefault(s => termSetId.Equals(s.Id, StringComparison.OrdinalIgnoreCase));

                if (set != null)
                {
                    return (group, set);
                }
            }

            return (null, null);
        }
    }
}
