using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using FieldModel = PnP.Core.Provisioning.Model.Field;
using TermGroupModel = PnP.Core.Provisioning.Model.TermGroup;
using TermModel = PnP.Core.Provisioning.Model.Term;
using TermSetModel = PnP.Core.Provisioning.Model.TermSet;

namespace PnP.Core.Provisioning.Test.Scenarios
{
    /// <summary>
    /// Scenario 5 - term groups, a taxonomy site column bound to the term set, and a list that uses
    /// it.
    /// </summary>
    /// <remarks>
    /// <para>The highest-risk item in the backlog, and the reason decision <b>D7</b> exists.
    /// <b>Every id in this template is supplied by the caller</b> - the group, the set and each term
    /// - and the whole point is that they survive. PnP Core's Graph-based term store cannot preserve
    /// a caller-supplied GUID, which is why spike S1 sent [T10](03-porting-backlog.md#t10) down the
    /// CSOM path.</para>
    /// <para><b>Why the ids matter beyond taxonomy.</b> A taxonomy site column binds to its term set
    /// by id, and a list item's taxonomy value stores the term's id. If the ids drift on apply, the
    /// column binds to nothing and every stored value is orphaned - and neither shows up as an
    /// error. So this scenario asserts the ids, then asserts the column resolved to them.</para>
    /// </remarks>
    [TestClass]
    public class Scenario05TaxonomyTests : ScenarioTestBase
    {
        private const string Prefix = "PnPCoreScenario5_";

        private static readonly Guid GroupId = Guid.NewGuid();
        private static readonly Guid TermSetId = Guid.NewGuid();
        private static readonly Guid FirstTermId = Guid.NewGuid();
        private static readonly Guid SecondTermId = Guid.NewGuid();

        private static readonly string GroupName = $"{Prefix}Group_{GroupId:N}".Substring(0, 40);
        private static readonly string TermSetName = $"{Prefix}Set";
        private static readonly string TaxonomyFieldId = "{6b1f2a3c-9d4e-4a5b-8c7d-2e1f0a9b8c05}";
        private static readonly string ListUrl = $"Lists/{Prefix}Tagged";

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Scenario")]
        [TestCategory("Taxonomy")]
        [Timeout(45 * 60 * 1000)]
        public async Task Scenario5_TermGroupsAndATaxonomyColumnBoundToThem()
        {
            try
            {
                await RunScenarioAsync("s5", BuildTemplate(), new[]
                {
                    ConfigurationHandler.Taxonomy,
                    ConfigurationHandler.Fields,
                    ConfigurationHandler.Lists,
                },
                AssertAsync,

                // Asked for explicitly, because the default is right and this scenario is the
                // exception. Extracting a site reports only that site's own term group - a site
                // extract that dragged in every tenant-wide group would produce a template nobody
                // could apply anywhere else. This scenario's group is tenant wide, so it has to say
                // so rather than the engine loosening its default.
                configuration => configuration.Taxonomy.IncludeAllTermGroups = true)
                .ConfigureAwait(false);
            }
            finally
            {
                // The term store is tenant wide - deleting the scenario's site does not touch it, so
                // the group has to go separately or it leaks. This is the one scenario whose
                // artefacts outlive its site.
                await CleanUpTermGroupAsync().ConfigureAwait(false);
            }
        }

        private static ProvisioningTemplate BuildTemplate()
        {
            var template = new ProvisioningTemplate { Id = "SCENARIO-5" };

            var first = new TermModel
            {
                Id = FirstTermId,
                Name = $"First_{Prefix}",
                Description = "Scenario 5 first term",
                CustomSortOrder = 1,
            };

            first.Labels.Add(new TermLabel
            {
                Value = $"Alias_{Prefix}",
                Language = CultureInfo.GetCultureInfo("en-US").LCID,
                IsDefaultForLanguage = false,
            });

            var second = new TermModel
            {
                Id = SecondTermId,
                Name = $"Second_{Prefix}",
                CustomSortOrder = 2,
            };

            var set = new TermSetModel
            {
                Id = TermSetId,
                Name = TermSetName,
                Description = "Scenario 5 term set",
                IsOpenForTermCreation = false,
            };

            set.Terms.Add(first);
            set.Terms.Add(second);

            var group = new TermGroupModel
            {
                Id = GroupId,
                Name = GroupName,
                Description = "Created by the PnP Core provisioning scenario tests",
            };

            group.TermSets.Add(set);
            template.TermGroups.Add(group);

            // A taxonomy column bound to the term set by id. The two GUIDs in the schema are the ones
            // the template declared above - if the term store did not honour them, this column binds
            // to a set that does not exist.
            template.SiteFields.Add(new FieldModel
            {
                SchemaXml =
                    $"<Field ID=\"{TaxonomyFieldId}\" Type=\"TaxonomyFieldType\" Name=\"{Prefix}Category\" " +
                    $"StaticName=\"{Prefix}Category\" DisplayName=\"Category\" Group=\"{Prefix}Group\" " +
                    "ShowField=\"Term1033\">" +
                    "<Customization><ArrayOfProperty>" +
                    "<Property><Name>SspId</Name><Value xmlns:q1=\"http://www.w3.org/2001/XMLSchema\" " +
                    "p4:type=\"q1:string\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">" +
                    // The default keywords term store, which is where ObjectTermGroups puts a group
                    // that names no store. A hardcoded id here would tie the scenario to one tenant.
                    "{keywordstermstoreid}</Value></Property>" +
                    "<Property><Name>TermSetId</Name><Value xmlns:q2=\"http://www.w3.org/2001/XMLSchema\" " +
                    "p4:type=\"q2:string\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">" +
                    $"{TermSetId}</Value></Property>" +
                    "</ArrayOfProperty></Customization>" +
                    "</Field>",
            });

            var list = new ListInstance
            {
                Title = $"{Prefix}Tagged",
                Url = ListUrl,
                TemplateType = (int)ListTemplateType.GenericList,
            };

            list.FieldRefs.Add(new FieldRef($"{Prefix}Category")
            {
                Id = Guid.Parse(TaxonomyFieldId),
            });

            template.Lists.Add(list);

            return template;
        }

        private static async Task AssertAsync(ProvisioningTemplate extracted, PnPContext site)
        {
            // ------------------------------------------------------------------------------------
            // The assertion the scenario exists for: the caller's GUIDs survived.
            // ------------------------------------------------------------------------------------
            ITermGroup group = await site.TermStore.Groups.GetByIdAsync(GroupId.ToString(),
                g => g.Id, g => g.Name).ConfigureAwait(false);

            Assert.IsNotNull(group,
                $"The term group was not created with the id the template gave it ({GroupId}). " +
                "This is decision D7 failing: a Graph-created group would have a server-assigned id.");

            Assert.AreEqual(GroupName, group.Name, "The term group has the wrong name.");

            await group.LoadAsync(g => g.Sets.QueryProperties(s => s.Id)).ConfigureAwait(false);

            ITermSet set = group.Sets.AsRequested().FirstOrDefault(s => s.Id == TermSetId.ToString());

            Assert.IsNotNull(set,
                $"The term set was not created with the id the template gave it ({TermSetId}).");

            await set.LoadAsync(s => s.Terms.QueryProperties(t => t.Id)).ConfigureAwait(false);

            List<string> termIds = set.Terms.AsRequested().Select(t => t.Id).ToList();

            Console.WriteLine($"Term ids: {string.Join(", ", termIds)}");

            CollectionAssert.Contains(termIds, FirstTermId.ToString(),
                $"The first term did not keep the id the template gave it ({FirstTermId}).");
            CollectionAssert.Contains(termIds, SecondTermId.ToString(),
                $"The second term did not keep the id the template gave it ({SecondTermId}).");

            // ------------------------------------------------------------------------------------
            // And the column bound to the set those ids belong to.
            // ------------------------------------------------------------------------------------
            await site.Web.LoadAsync(w => w.Fields.QueryProperties(f => f.InternalName, f => f.SchemaXml))
                .ConfigureAwait(false);

            IField column = site.Web.Fields.AsRequested()
                .FirstOrDefault(f => f.InternalName == $"{Prefix}Category");

            Assert.IsNotNull(column, "The taxonomy site column was not created.");

            string boundSetId = TermSetIdOf(column.SchemaXml);

            Console.WriteLine($"Column bound to term set: {boundSetId}");

            Assert.AreEqual(TermSetId.ToString(), boundSetId?.ToLowerInvariant(),
                "The taxonomy column is not bound to the term set the template created. A column " +
                "bound to a set that does not exist looks identical until someone tries to tag with it.");

            // ------------------------------------------------------------------------------------
            // The list took the column, and the extract found the group.
            // ------------------------------------------------------------------------------------
            await site.Web.LoadAsync(w => w.Lists.QueryProperties(l => l.Title,
                l => l.Fields.QueryProperties(f => f.InternalName))).ConfigureAwait(false);

            IList list = site.Web.Lists.AsRequested().FirstOrDefault(l => l.Title == $"{Prefix}Tagged");

            Assert.IsNotNull(list, "The list was not created.");
            Assert.IsTrue(list.Fields.AsRequested().Any(f => f.InternalName == $"{Prefix}Category"),
                "The taxonomy column did not reach the list that references it.");

            Assert.IsTrue(extracted.TermGroups.Any(g => g.Id == GroupId),
                $"The extract did not report the term group by its id. " +
                $"Found: {string.Join(", ", extracted.TermGroups.Select(g => g.Id))}");
        }

        /// <summary>
        /// Reads the <c>TermSetId</c> a taxonomy column is bound to out of its schema.
        /// </summary>
        /// <remarks>
        /// It lives in a <c>Customization</c> property bag rather than an attribute, which is why
        /// this is a search through elements rather than an attribute read.
        /// </remarks>
        private static string TermSetIdOf(string schemaXml)
        {
            if (string.IsNullOrEmpty(schemaXml))
            {
                return null;
            }

            foreach (XElement property in XElement.Parse(schemaXml).Descendants()
                .Where(e => e.Name.LocalName == "Property"))
            {
                XElement name = property.Elements().FirstOrDefault(e => e.Name.LocalName == "Name");

                if (name?.Value == "TermSetId")
                {
                    return property.Elements().FirstOrDefault(e => e.Name.LocalName == "Value")?.Value;
                }
            }

            return null;
        }

        private static async Task CleanUpTermGroupAsync()
        {
            try
            {
                using (PnPContext context = await GetContextAsync(4).ConfigureAwait(false))
                {
                    await DeleteTermGroupDeepAsync(context, GroupId.ToString()).ConfigureAwait(false);
                    Console.WriteLine($"Deleted term group {GroupId}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"COULD NOT DELETE TERM GROUP {GroupId}: {Describe(ex)}");
            }
        }
    }
}
