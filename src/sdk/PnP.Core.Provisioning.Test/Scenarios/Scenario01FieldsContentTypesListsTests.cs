using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using ContentTypeModel = PnP.Core.Provisioning.Model.ContentType;
using FieldModel = PnP.Core.Provisioning.Model.Field;

namespace PnP.Core.Provisioning.Test.Scenarios
{
    /// <summary>
    /// Scenario 1 - fields, content types, lists and a <b>cross-list lookup</b>.
    /// </summary>
    /// <remarks>
    /// <para>Described in the test strategy as the single most important test in the suite, and the
    /// reason is one line of the template: <b>the lookup column points at a list defined later in
    /// the same template.</b> Nothing but the three-pass ordering makes that work - pass one creates
    /// the lists and the non-lookup columns, pass two creates the lookups now that their targets
    /// exist, pass three applies the list settings that reference them.</para>
    /// <para>It is worth being precise about what would happen without the ordering, because it is
    /// not a crash: SharePoint accepts a lookup whose <c>List</c> attribute names a url it cannot
    /// resolve and creates a column pointing at nothing. The apply would report success. So this
    /// scenario asserts on the <b>resolved lookup target</b>, not on the column's existence.</para>
    /// </remarks>
    [TestClass]
    public class Scenario01FieldsContentTypesListsTests : ScenarioTestBase
    {
        private const string LookupFieldId = "{4f1d0b7e-3a52-4a1e-9a1c-9a9d1b7c5e01}";
        private const string TextFieldId = "{7a2c5f19-9e64-4d3b-8c2a-1f4b6d8e2a02}";
        private const string ContentTypeId = "0x0100A9E1B4C25D3E4F5A8B7C6D5E4F3A2B1C";

        private static readonly string TargetListUrl = $"Lists/{ScenarioPrefix}Target";
        private static readonly string SourceListUrl = $"Lists/{ScenarioPrefix}Source";

        private const string ScenarioPrefix = "PnPCoreScenario1_";

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Scenario")]
        [Timeout(45 * 60 * 1000)]
        public async Task Scenario1_FieldsContentTypesListsAndACrossListLookup()
        {
            await RunScenarioAsync("s1", BuildTemplate(), new[]
            {
                ConfigurationHandler.Fields,
                ConfigurationHandler.ContentTypes,
                ConfigurationHandler.Lists,
            },
            AssertAsync).ConfigureAwait(false);
        }

        /// <summary>
        /// The template. <b>The source list is declared before the list its lookup points at.</b>
        /// </summary>
        /// <remarks>
        /// Declaration order is deliberately hostile: a single-pass engine reading this top to bottom
        /// meets the lookup before the list it targets. That is exactly the case the three passes
        /// exist for, and putting the target first would make the scenario pass for the wrong reason.
        /// </remarks>
        private static ProvisioningTemplate BuildTemplate()
        {
            var template = new ProvisioningTemplate { Id = "SCENARIO-1" };

            // A plain site column, used through a content type.
            template.SiteFields.Add(new FieldModel
            {
                SchemaXml = $"<Field ID=\"{TextFieldId}\" Type=\"Text\" Name=\"{ScenarioPrefix}Code\" " +
                    $"StaticName=\"{ScenarioPrefix}Code\" DisplayName=\"Code\" Group=\"{ScenarioPrefix}Group\" />",
            });

            // The lookup, pointing at a list that does not exist yet.
            template.SiteFields.Add(new FieldModel
            {
                SchemaXml = $"<Field ID=\"{LookupFieldId}\" Type=\"Lookup\" Name=\"{ScenarioPrefix}Target\" " +
                    $"StaticName=\"{ScenarioPrefix}Target\" DisplayName=\"Target\" " +
                    $"Group=\"{ScenarioPrefix}Group\" List=\"{TargetListUrl}\" ShowField=\"Title\" />",
            });

            var contentType = new ContentTypeModel
            {
                Id = ContentTypeId,
                Name = $"{ScenarioPrefix}Item",
                Description = "Scenario 1 content type",
                Group = $"{ScenarioPrefix}Group",
            };

            contentType.FieldRefs.Add(new FieldRef($"{ScenarioPrefix}Code")
            {
                Id = Guid.Parse(TextFieldId),
            });

            template.ContentTypes.Add(contentType);

            // Declared first, and it is the one carrying the lookup.
            var source = new ListInstance
            {
                Title = $"{ScenarioPrefix}Source",
                Url = SourceListUrl,
                TemplateType = (int)ListTemplateType.GenericList,
                ContentTypesEnabled = true,
            };

            source.ContentTypeBindings.Add(new ContentTypeBinding { ContentTypeId = ContentTypeId, Default = true });

            source.FieldRefs.Add(new FieldRef($"{ScenarioPrefix}Target")
            {
                Id = Guid.Parse(LookupFieldId),
            });

            template.Lists.Add(source);

            // Declared second, and it is what the lookup above points at.
            template.Lists.Add(new ListInstance
            {
                Title = $"{ScenarioPrefix}Target",
                Url = TargetListUrl,
                TemplateType = (int)ListTemplateType.GenericList,
            });

            return template;
        }

        private static async Task AssertAsync(ProvisioningTemplate extracted, PnPContext site)
        {
            // ------------------------------------------------------------------------------------
            // The assertion the scenario exists for: the lookup resolved to the right list.
            // ------------------------------------------------------------------------------------
            await site.Web.LoadAsync(w => w.Lists.QueryProperties(l => l.Id, l => l.Title))
                .ConfigureAwait(false);

            IList sourceList = site.Web.Lists.AsRequested()
                .FirstOrDefault(l => l.Title == $"{ScenarioPrefix}Source");
            IList targetList = site.Web.Lists.AsRequested()
                .FirstOrDefault(l => l.Title == $"{ScenarioPrefix}Target");

            Assert.IsNotNull(sourceList, "The source list was not created.");
            Assert.IsNotNull(targetList, "The target list was not created.");

            // SchemaXml is asked for here, with the rest. A field taken out of a collection that was
            // materialised without it does not gain it from a later LoadAsync - the collection's
            // items are built once, and asking again does not top them up. That behaviour cost
            // several rounds in phase 6 and it is the same trap here.
            await sourceList.LoadAsync(l => l.Fields.QueryProperties(
                f => f.InternalName, f => f.TypeAsString, f => f.SchemaXml)).ConfigureAwait(false);

            IField lookup = sourceList.Fields.AsRequested()
                .FirstOrDefault(f => f.InternalName == $"{ScenarioPrefix}Target");

            Assert.IsNotNull(lookup, $"The lookup column is not on the source list.");

            // Read as XML because a lookup's target is what matters and it lives in the schema. A
            // lookup pointing at nothing has the same InternalName and the same TypeAsString as one
            // pointing at the right list - the List attribute is the only thing that differs.
            string listAttribute = (string)XElement.Parse(lookup.SchemaXml).Attribute("List");

            Console.WriteLine($"Lookup List attribute: {listAttribute}");
            Console.WriteLine($"Target list id:        {targetList.Id}");

            Assert.IsFalse(string.IsNullOrEmpty(listAttribute),
                "The lookup column has no List attribute, so it points at nothing - which is what a " +
                "single-pass apply produces, and SharePoint reports as success.");

            Assert.AreEqual(targetList.Id, Guid.Parse(listAttribute.Trim('{', '}')),
                "The lookup column does not point at the list the template named. This is the " +
                "three-pass ordering failing: the lookup was created before its target list existed.");

            // ------------------------------------------------------------------------------------
            // The rest of the template landed too.
            // ------------------------------------------------------------------------------------
            await site.Web.LoadAsync(w => w.ContentTypes.QueryProperties(c => c.StringId, c => c.Name))
                .ConfigureAwait(false);

            Assert.IsTrue(site.Web.ContentTypes.AsRequested().Any(c => c.StringId == ContentTypeId),
                "The content type was not created.");

            Assert.IsTrue(sourceList.Fields.AsRequested().Any(f => f.InternalName == $"{ScenarioPrefix}Code"),
                "The content type's field did not reach the list that binds it.");

            // ------------------------------------------------------------------------------------
            // And the extract found them.
            // ------------------------------------------------------------------------------------
            Assert.IsTrue(extracted.Lists.Any(l => l.Title == $"{ScenarioPrefix}Source"),
                "The extract did not report the source list.");
            Assert.IsTrue(extracted.Lists.Any(l => l.Title == $"{ScenarioPrefix}Target"),
                "The extract did not report the target list.");

            List<string> extractedFieldNames = extracted.SiteFields
                .Select(f => (string)XElement.Parse(f.SchemaXml).Attribute("Name"))
                .ToList();

            Assert.IsTrue(extractedFieldNames.Contains($"{ScenarioPrefix}Target"),
                $"The extract did not report the lookup site column. Found: {string.Join(", ", extractedFieldNames)}");
        }
    }
}
