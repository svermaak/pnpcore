using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Providers;
using PnP.Core.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

// The provisioning model declares its own File and Directory types (template artefacts, not IO), so the
// System.IO ones have to be named explicitly - the same disambiguation the PnP Framework sample app needs.
using File = System.IO.File;
using Path = System.IO.Path;

namespace PnP.Core.Provisioning.Test.Offline.Providers
{
    /// <summary>
    /// The migration phase 2 exit gate: every supported provisioning schema version must deserialize
    /// into the domain model, serialize back out, and survive the trip intact.
    /// </summary>
    /// <remarks>
    /// These tests need no SharePoint connection - they exercise the XSD generated schema classes, the
    /// reflection driven serializer/resolver layer and the formatters in isolation. That covers the
    /// large majority of the ported code by volume, which is why this gate is worth having before any
    /// object handler is written.
    ///
    /// Exact XML equality is deliberately NOT asserted. Serializing a template legitimately reorders
    /// attributes, materializes defaults and drops empty optional elements - PnP Framework ships
    /// "-OUT.xml" variants of some fixtures for exactly this reason. What must hold is that the *model*
    /// survives: template -> XML -> template produces an equivalent template.
    /// </remarks>
    [TestClass]
    public class SchemaRoundTripTests
    {
        /// <summary>
        /// Every schema version the engine supports, with the sample template that exercises it.
        /// </summary>
        public static IEnumerable<object[]> SchemaVersions()
        {
            yield return new object[] { "ProvisioningSchema-2019-03-FullSample-01.xml", XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2019_03 };
            yield return new object[] { "ProvisioningSchema-2019-09-FullSample-01.xml", XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2019_09 };
            yield return new object[] { "ProvisioningSchema-2020-02-FullSample-01.xml", XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2020_02 };
            yield return new object[] { "ProvisioningSchema-2021-03-FullSample-01.xml", XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2021_03 };
            yield return new object[] { "ProvisioningSchema-2022-09-FullSample-01.xml", XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2022_09 };
        }

        private static string TemplatesFolder => Path.Combine(AppContext.BaseDirectory, "TestAssets", "Templates");

        private static Stream OpenFixture(string fileName)
        {
            string path = Path.Combine(TemplatesFolder, fileName);
            Assert.IsTrue(File.Exists(path), $"Fixture not found: {path}");
            return File.OpenRead(path);
        }

        [TestMethod]
        [TestCategory("Offline")]
        [DynamicData(nameof(SchemaVersions), DynamicDataSourceType.Method)]
        public void Formatter_IsResolvedFromTheTemplateNamespace(string fileName, string namespaceUri)
        {
            _ = fileName;

            ITemplateFormatter formatter = XMLPnPSchemaFormatter.GetSpecificFormatter(namespaceUri);

            Assert.IsNotNull(formatter);
            Assert.AreEqual(namespaceUri, ((IXMLSchemaFormatter)formatter).NamespaceUri,
                "The formatter selected for a namespace must serialize into that same namespace.");
        }

        [TestMethod]
        [TestCategory("Offline")]
        [DynamicData(nameof(SchemaVersions), DynamicDataSourceType.Method)]
        public void Deserialize_ProducesAPopulatedTemplate(string fileName, string namespaceUri)
        {
            ITemplateFormatter formatter = XMLPnPSchemaFormatter.GetSpecificFormatter(namespaceUri);

            ProvisioningTemplate template;
            using (Stream fixture = OpenFixture(fileName))
            {
                template = formatter.ToProvisioningTemplate(fixture);
            }

            Assert.IsNotNull(template, "Deserialization returned no template");

            // The full samples are deliberately broad; if the serializer/resolver layer were mis-wired
            // these would come back empty rather than throwing.
            Assert.IsFalse(string.IsNullOrEmpty(template.Id), "Template Id was not deserialized");
            Assert.IsTrue(template.SiteFields.Count > 0, "Site fields were not deserialized");
            Assert.IsTrue(template.ContentTypes.Count > 0, "Content types were not deserialized");
            Assert.IsTrue(template.Lists.Count > 0, "Lists were not deserialized");
        }

        [TestMethod]
        [TestCategory("Offline")]
        [DynamicData(nameof(SchemaVersions), DynamicDataSourceType.Method)]
        public void Serialize_ProducesXmlInTheExpectedNamespace(string fileName, string namespaceUri)
        {
            ITemplateFormatter formatter = XMLPnPSchemaFormatter.GetSpecificFormatter(namespaceUri);

            ProvisioningTemplate template;
            using (Stream fixture = OpenFixture(fileName))
            {
                template = formatter.ToProvisioningTemplate(fixture);
            }

            using (Stream serialized = formatter.ToFormattedTemplate(template))
            {
                XDocument document = XDocument.Load(serialized);

                Assert.IsNotNull(document.Root);
                Assert.AreEqual(namespaceUri, document.Root.Name.NamespaceName,
                    "Serialized output landed in the wrong schema namespace.");
            }
        }

        [TestMethod]
        [TestCategory("Offline")]
        [DynamicData(nameof(SchemaVersions), DynamicDataSourceType.Method)]
        public void Serialize_OutputValidatesAgainstItsSchema(string fileName, string namespaceUri)
        {
            ITemplateFormatter formatter = XMLPnPSchemaFormatter.GetSpecificFormatter(namespaceUri);

            ProvisioningTemplate template;
            using (Stream fixture = OpenFixture(fileName))
            {
                template = formatter.ToProvisioningTemplate(fixture);
            }

            using (Stream serialized = formatter.ToFormattedTemplate(template))
            {
                // Validation runs against the embedded .xsd, so this also proves the schemas are
                // correctly wired up as embedded resources under their new PnP.Core resource names.
                ValidationResult validation = ((ITemplateFormatterWithValidation)formatter).GetValidationResults(serialized);

                string failures = validation.Exceptions == null
                    ? string.Empty
                    : string.Join(Environment.NewLine, validation.Exceptions.Select(e => e.Message));

                Assert.IsTrue(validation.IsValid, $"Serialized template did not validate against its schema:{Environment.NewLine}{failures}");
            }
        }

        [TestMethod]
        [TestCategory("Offline")]
        [DynamicData(nameof(SchemaVersions), DynamicDataSourceType.Method)]
        public void RoundTrip_ModelSurvivesSerializeAndDeserialize(string fileName, string namespaceUri)
        {
            ITemplateFormatter formatter = XMLPnPSchemaFormatter.GetSpecificFormatter(namespaceUri);

            ProvisioningTemplate original;
            using (Stream fixture = OpenFixture(fileName))
            {
                original = formatter.ToProvisioningTemplate(fixture);
            }

            ProvisioningTemplate roundTripped;
            using (Stream serialized = formatter.ToFormattedTemplate(original))
            {
                roundTripped = formatter.ToProvisioningTemplate(serialized);
            }

            Assert.IsNotNull(roundTripped);

            // Compare the shape of the template rather than object identity. A mis-wired resolver
            // typically shows up as a collection that empties on the way out.
            Assert.AreEqual(original.Id, roundTripped.Id, nameof(original.Id));
            Assert.AreEqual(original.SiteFields.Count, roundTripped.SiteFields.Count, nameof(original.SiteFields));
            Assert.AreEqual(original.ContentTypes.Count, roundTripped.ContentTypes.Count, nameof(original.ContentTypes));
            Assert.AreEqual(original.Lists.Count, roundTripped.Lists.Count, nameof(original.Lists));
            Assert.AreEqual(original.Features?.SiteFeatures?.Count, roundTripped.Features?.SiteFeatures?.Count, "SiteFeatures");
            Assert.AreEqual(original.Features?.WebFeatures?.Count, roundTripped.Features?.WebFeatures?.Count, "WebFeatures");
            Assert.AreEqual(original.CustomActions?.SiteCustomActions?.Count, roundTripped.CustomActions?.SiteCustomActions?.Count, "SiteCustomActions");
            Assert.AreEqual(original.CustomActions?.WebCustomActions?.Count, roundTripped.CustomActions?.WebCustomActions?.Count, "WebCustomActions");
            Assert.AreEqual(original.PropertyBagEntries.Count, roundTripped.PropertyBagEntries.Count, nameof(original.PropertyBagEntries));
            Assert.AreEqual(original.Files.Count, roundTripped.Files.Count, nameof(original.Files));
            Assert.AreEqual(original.TermGroups.Count, roundTripped.TermGroups.Count, nameof(original.TermGroups));
            Assert.AreEqual(original.Security?.SiteGroups?.Count, roundTripped.Security?.SiteGroups?.Count, "SiteGroups");
            Assert.AreEqual(original.Security?.SiteSecurityPermissions?.RoleDefinitions?.Count,
                            roundTripped.Security?.SiteSecurityPermissions?.RoleDefinitions?.Count, "RoleDefinitions");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void ToXml_UsesTheLatestSchemaByDefault()
        {
            // ProvisioningTemplate.ToXML() was held back during the model port because it depends on the
            // formatter chain; this guards its restoration.
            ITemplateFormatter formatter = XMLPnPSchemaFormatter.GetSpecificFormatter(XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2022_09);

            ProvisioningTemplate template;
            using (Stream fixture = OpenFixture("ProvisioningSchema-2022-09-FullSample-01.xml"))
            {
                template = formatter.ToProvisioningTemplate(fixture);
            }

            string xml = template.ToXML();

            Assert.IsFalse(string.IsNullOrEmpty(xml));
            Assert.IsTrue(xml.Contains(XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2022_09),
                "ToXML() should emit the latest schema namespace.");
        }
    }
}
