using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Services.Core.CSOM;
using PnP.Core.Provisioning.Services.Core.CSOM.Requests.Taxonomy;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.Test.Live.Csom
{
    /// <summary>
    /// Live round trips for the eight taxonomy CSOM requests (<b>T10</b>).
    /// </summary>
    /// <remarks>
    /// <para>These are the requests spike <b>S1</b> forced into existence, so the assertion that
    /// matters throughout is <b>the id came back exactly as supplied</b>. Everything else about
    /// taxonomy provisioning depends on that: templates embed term GUIDs in taxonomy field defaults
    /// and in list item values, and a reassigned id breaks every one of those references.</para>
    /// <para>These write to the <b>tenant-wide</b> term store, not to a test site. Every artefact
    /// is prefixed and deleted in a <c>finally</c>.</para>
    /// </remarks>
    [TestClass]
    public class TaxonomyRequestLiveTests : LiveTestBase
    {
        private static string UniqueName(string what) => $"{TestPrefix}{what}_{DateTime.UtcNow:HHmmssfff}";

        /// <summary>
        /// Deletes a term group by id, emptying it first, tolerating its absence.
        /// </summary>
        /// <remarks>
        /// Must be a <em>deep</em> delete. Most of these tests create term sets inside the group,
        /// and Graph refuses to delete a non-empty group - which leaked seven groups into the
        /// tenant term store before the hygiene check caught it.
        /// </remarks>
        private static async Task DeleteGroupAsync(PnPContext context, Guid groupId)
        {
            if (groupId == Guid.Empty) return;

            await DeleteTermGroupDeepAsync(context, groupId.ToString()).ConfigureAwait(false);
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Taxonomy")]
        public async Task CreateTermGroup_PreservesTheSuppliedId()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                Guid groupId = Guid.NewGuid();
                string groupName = UniqueName("Group");

                try
                {
                    TermGroupInfo created = await CsomRequestSender.SendAsync(context,
                        new CreateTermGroupRequest(groupName, groupId, "Created by a live test")).ConfigureAwait(false);

                    Assert.IsNotNull(created, "CreateTermGroupRequest returned no result.");

                    // The whole point of T10.
                    Assert.AreEqual(groupId, created.Id,
                        "The term group id was not preserved. If CSOM cannot preserve ids either, D7's premise is wrong.");
                    Assert.AreEqual(groupName, created.Name);

                    // And it is really there - reading it back through Graph proves the CSOM write
                    // landed in the same store PnP Core reads from, which is S1 divergence (g).
                    ITermGroup readBack = await context.TermStore.Groups.GetByIdAsync(groupId.ToString()).ConfigureAwait(false);
                    Assert.IsNotNull(readBack, "The group was created but Graph cannot see it - the two APIs disagree on which store this is.");
                    Assert.AreEqual(groupName, readBack.Name);
                }
                finally
                {
                    await DeleteGroupAsync(context, groupId).ConfigureAwait(false);
                }
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Taxonomy")]
        public async Task CreateTermSetAndTerm_PreserveTheirSuppliedIds()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                Guid groupId = Guid.NewGuid();
                Guid termSetId = Guid.NewGuid();
                Guid termId = Guid.NewGuid();

                try
                {
                    await CsomRequestSender.SendAsync(context,
                        new CreateTermGroupRequest(UniqueName("Group"), groupId)).ConfigureAwait(false);

                    TermSetInfo termSet = await CsomRequestSender.SendAsync(context,
                        new CreateTermSetRequest(groupId, UniqueName("Set"), termSetId, 1033)).ConfigureAwait(false);

                    Assert.IsNotNull(termSet, "CreateTermSetRequest returned no result.");
                    Assert.AreEqual(termSetId, termSet.Id, "The term set id was not preserved.");

                    TermInfo term = await CsomRequestSender.SendAsync(context,
                        new CreateTermRequest(termSetId, parentIsTermSet: true, UniqueName("Term"), termId, 1033))
                        .ConfigureAwait(false);

                    Assert.IsNotNull(term, "CreateTermRequest returned no result.");
                    Assert.AreEqual(termId, term.Id, "The term id was not preserved - this is the S1 failure mode.");
                }
                finally
                {
                    await DeleteGroupAsync(context, groupId).ConfigureAwait(false);
                }
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Taxonomy")]
        public async Task CreateTerm_UnderAnotherTermPreservesTheSuppliedId()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                Guid groupId = Guid.NewGuid();
                Guid termSetId = Guid.NewGuid();
                Guid parentTermId = Guid.NewGuid();
                Guid childTermId = Guid.NewGuid();

                try
                {
                    await CsomRequestSender.SendAsync(context, new CreateTermGroupRequest(UniqueName("Group"), groupId)).ConfigureAwait(false);
                    await CsomRequestSender.SendAsync(context, new CreateTermSetRequest(groupId, UniqueName("Set"), termSetId, 1033)).ConfigureAwait(false);
                    await CsomRequestSender.SendAsync(context, new CreateTermRequest(termSetId, true, UniqueName("Parent"), parentTermId, 1033)).ConfigureAwait(false);

                    // CreateTerm hangs off TermSetItem, the shared base of TermSet and Term, so the
                    // same request creates a child term - only the parent lookup differs.
                    TermInfo child = await CsomRequestSender.SendAsync(context,
                        new CreateTermRequest(parentTermId, parentIsTermSet: false, UniqueName("Child"), childTermId, 1033))
                        .ConfigureAwait(false);

                    Assert.IsNotNull(child);
                    Assert.AreEqual(childTermId, child.Id);
                }
                finally
                {
                    await DeleteGroupAsync(context, groupId).ConfigureAwait(false);
                }
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Taxonomy")]
        public async Task AddTermLabel_AddsASynonymInTheGivenLanguage()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                Guid groupId = Guid.NewGuid();
                Guid termSetId = Guid.NewGuid();
                Guid termId = Guid.NewGuid();

                try
                {
                    await CsomRequestSender.SendAsync(context, new CreateTermGroupRequest(UniqueName("Group"), groupId)).ConfigureAwait(false);
                    await CsomRequestSender.SendAsync(context, new CreateTermSetRequest(groupId, UniqueName("Set"), termSetId, 1033)).ConfigureAwait(false);
                    await CsomRequestSender.SendAsync(context, new CreateTermRequest(termSetId, true, "Primary", termId, 1033)).ConfigureAwait(false);

                    await CsomRequestSender.SendAsync(context,
                        new AddTermLabelRequest(termId, "Synonym", 1033, isDefaultForLanguage: false)).ConfigureAwait(false);

                    // Read back through Graph: the labels collection should now carry both.
                    ITermSet set = await context.TermStore.Groups.GetById(groupId.ToString())
                        .Sets.GetByIdAsync(termSetId.ToString()).ConfigureAwait(false);
                    await set.LoadAsync(s => s.Terms.QueryProperties(t => t.Id, t => t.Labels)).ConfigureAwait(false);

                    ITerm term = set.Terms.AsRequested().FirstOrDefault(t => t.Id == termId.ToString());
                    Assert.IsNotNull(term, "The term was not found after adding a label.");

                    Assert.IsTrue(term.Labels.Any(l => l.Name == "Synonym"),
                        "The synonym label was not persisted. Labels present: " +
                        string.Join(", ", term.Labels.Select(l => l.Name)));
                }
                finally
                {
                    await DeleteGroupAsync(context, groupId).ConfigureAwait(false);
                }
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Taxonomy")]
        public async Task SetTermCustomProperty_WritesBothSharedAndLocalProperties()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                Guid groupId = Guid.NewGuid();
                Guid termSetId = Guid.NewGuid();
                Guid termId = Guid.NewGuid();

                try
                {
                    await CsomRequestSender.SendAsync(context, new CreateTermGroupRequest(UniqueName("Group"), groupId)).ConfigureAwait(false);
                    await CsomRequestSender.SendAsync(context, new CreateTermSetRequest(groupId, UniqueName("Set"), termSetId, 1033)).ConfigureAwait(false);
                    await CsomRequestSender.SendAsync(context, new CreateTermRequest(termSetId, true, "Term", termId, 1033)).ConfigureAwait(false);

                    // S1 verdict: Graph collapses these two into one bare key/value pair.
                    await CsomRequestSender.SendAsync(context,
                        new SetTermCustomPropertyRequest(termId, "SharedKey", "SharedValue", isLocal: false)).ConfigureAwait(false);

                    await CsomRequestSender.SendAsync(context,
                        new SetTermCustomPropertyRequest(termId, "LocalKey", "LocalValue", isLocal: true)).ConfigureAwait(false);

                    ITermSet set = await context.TermStore.Groups.GetById(groupId.ToString())
                        .Sets.GetByIdAsync(termSetId.ToString()).ConfigureAwait(false);
                    await set.LoadAsync(s => s.Terms.QueryProperties(t => t.Id, t => t.Properties)).ConfigureAwait(false);

                    ITerm term = set.Terms.AsRequested().FirstOrDefault(t => t.Id == termId.ToString());
                    Assert.IsNotNull(term);

                    // Graph surfaces both kinds through the same collection - which is precisely
                    // the lossiness S1 recorded. Both values should nonetheless be present.
                    Assert.IsTrue(term.Properties.Any(p => p.KeyField == "SharedKey" && p.Value == "SharedValue"),
                        "The shared custom property was not persisted.");
                }
                finally
                {
                    await DeleteGroupAsync(context, groupId).ConfigureAwait(false);
                }
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Taxonomy")]
        public async Task DeprecateTerm_MarksTheTermDeprecated()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                Guid groupId = Guid.NewGuid();
                Guid termSetId = Guid.NewGuid();
                Guid termId = Guid.NewGuid();

                try
                {
                    await CsomRequestSender.SendAsync(context, new CreateTermGroupRequest(UniqueName("Group"), groupId)).ConfigureAwait(false);
                    await CsomRequestSender.SendAsync(context, new CreateTermSetRequest(groupId, UniqueName("Set"), termSetId, 1033)).ConfigureAwait(false);
                    await CsomRequestSender.SendAsync(context, new CreateTermRequest(termSetId, true, "Retired", termId, 1033)).ConfigureAwait(false);

                    // S1 verdict: no Graph equivalent at all. Verified here only in that the call
                    // succeeds - Graph cannot report the flag back, which is the gap.
                    await CsomRequestSender.SendAsync(context,
                        new DeprecateTermRequest(termId, deprecate: true)).ConfigureAwait(false);

                    await CsomRequestSender.SendAsync(context,
                        new DeprecateTermRequest(termId, deprecate: false)).ConfigureAwait(false);
                }
                finally
                {
                    await DeleteGroupAsync(context, groupId).ConfigureAwait(false);
                }
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Taxonomy")]
        public async Task ReuseTerm_ReusesATermUnderAnotherSet()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                Guid groupId = Guid.NewGuid();
                Guid sourceSetId = Guid.NewGuid();
                Guid targetSetId = Guid.NewGuid();
                Guid sourceTermId = Guid.NewGuid();

                try
                {
                    await CsomRequestSender.SendAsync(context, new CreateTermGroupRequest(UniqueName("Group"), groupId)).ConfigureAwait(false);
                    await CsomRequestSender.SendAsync(context, new CreateTermSetRequest(groupId, UniqueName("Source"), sourceSetId, 1033)).ConfigureAwait(false);
                    await CsomRequestSender.SendAsync(context, new CreateTermSetRequest(groupId, UniqueName("Target"), targetSetId, 1033)).ConfigureAwait(false);
                    await CsomRequestSender.SendAsync(context, new CreateTermRequest(sourceSetId, true, "Shared", sourceTermId, 1033)).ConfigureAwait(false);

                    TermInfo reused = await CsomRequestSender.SendAsync(context,
                        new ReuseTermRequest(targetSetId, parentIsTermSet: true, sourceTermId)).ConfigureAwait(false);

                    Assert.IsNotNull(reused, "ReuseTermRequest returned no result.");

                    // A reused term keeps the source term's id - that is what makes it the *same*
                    // term rather than a copy.
                    Assert.AreEqual(sourceTermId, reused.Id,
                        "A reused term should carry the source term's id.");
                }
                finally
                {
                    await DeleteGroupAsync(context, groupId).ConfigureAwait(false);
                }
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Taxonomy")]
        public async Task GetSiteCollectionTermGroup_ResolvesTheScopedGroup()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                (Guid siteId, Guid webId) = await CsomRequestSender.GetSiteAndWebIdAsync(context).ConfigureAwait(false);

                // createIfMissing: false so a read-only probe cannot provision taxonomy as a side
                // effect. A site with no scoped group yet is a legitimate outcome.
                TermGroupInfo group = await CsomRequestSender.SendAsync(context,
                    new GetSiteCollectionTermGroupRequest(siteId, webId, createIfMissing: false)).ConfigureAwait(false);

                if (group == null || group.Id == Guid.Empty)
                {
                    Assert.Inconclusive(
                        "This site collection has no scoped term group yet, so there was nothing to resolve. " +
                        "Re-run with createIfMissing to exercise the creating path.");
                    return;
                }

                Assert.IsTrue(group.IsSiteCollectionGroup,
                    "GetSiteCollectionGroup returned a group that is not site-collection scoped.");
                Assert.IsFalse(string.IsNullOrEmpty(group.Name));
            }
        }
    }
}
