using PnP.Core.Services.Core.CSOM;
using PnP.Core.Services.Core.CSOM.QueryAction;
using PnP.Core.Services.Core.CSOM.Requests;
using PnP.Core.Services.Core.CSOM.Utils;
using System;
using System.Collections.Generic;

namespace PnP.Core.Provisioning.Services.Core.CSOM.Requests.Taxonomy
{
    /// <summary>
    /// Sets the explicit child ordering of a term set or a term.
    /// </summary>
    /// <remarks>
    /// <para>One of the four divergences that forced the CSOM-first taxonomy decision
    /// (<b>D7</b>): <b>Graph has no representation of custom sort order at all</b>. A term store
    /// browsed through Graph is always alphabetical.</para>
    /// <para>The order is a property of the <em>parent</em>, not of each child - a colon-separated
    /// list of child term ids. So it is written once, after all the siblings exist, and a term whose
    /// creation failed simply does not appear in it.</para>
    /// <para>An empty string is meaningful: it clears the custom order and returns the parent to
    /// alphabetical.</para>
    /// </remarks>
    internal sealed class SetCustomSortOrderRequest : TaxonomyRequestBase, IRequest<object>
    {
        private readonly Guid parentId;
        private readonly bool parentIsTermSet;
        private readonly string customSortOrder;

        /// <summary>
        /// Creates the request.
        /// </summary>
        /// <param name="parentId">The term set or term whose children are being ordered</param>
        /// <param name="parentIsTermSet">Whether <paramref name="parentId"/> names a term set</param>
        /// <param name="customSortOrder">Colon-separated child term ids, or empty to clear</param>
        internal SetCustomSortOrderRequest(Guid parentId, bool parentIsTermSet, string customSortOrder)
        {
            if (parentId == Guid.Empty)
            {
                throw new ArgumentException("A parent id is required.", nameof(parentId));
            }

            this.parentId = parentId;
            this.parentIsTermSet = parentIsTermSet;
            this.customSortOrder = customSortOrder ?? string.Empty;
        }

        public object Result { get; private set; }

        public List<ActionObjectPath> GetRequest(IIdProvider idProvider)
        {
            var paths = new List<ActionObjectPath>();

            int termStoreId = EmitTermStorePrologue(idProvider, paths);

            int parentPathId = parentIsTermSet
                ? EmitGetTermSet(idProvider, paths, termStoreId, parentId)
                : EmitGetTerm(idProvider, paths, termStoreId, parentId);

            paths.Add(new ActionObjectPath
            {
                Action = new SetPropertyAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = parentPathId.ToString(),
                    Name = "CustomSortOrder",
                    SetParameter = new Parameter { Type = "String", Value = customSortOrder },
                },
            });

            // As everywhere in the term store, the change is staged until CommitAll.
            EmitCommitAll(idProvider, paths, termStoreId);

            return paths;
        }

        public void ProcessResponse(string response)
        {
        }
    }
}
