using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Provisioning.BaseTemplates;
using PnP.Core.Provisioning.Model;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace PnP.Core.Provisioning.Test.Offline.BaseTemplates
{
    /// <summary>
    /// Guards the 22 embedded out-of-the-box site templates.
    /// </summary>
    /// <remarks>
    /// These are 4.3 MB of XML wired up as embedded resources, and a mistake in that wiring is
    /// silent: <c>GetManifestResourceStream</c> returns null and extraction quietly reports every
    /// list SharePoint creates by default as if the template had asked for it. Loading each one
    /// and asserting it deserializes catches both the resource naming and any schema drift.
    /// </remarks>
    [TestClass]
    public class BaseTemplateManagerTests
    {
        /// <summary>
        /// Every base template that ships, as (web template, configuration) - i.e. the two halves
        /// of the resource name <c>{webTemplate}{configuration}Template.xml</c>.
        /// </summary>
        public static IEnumerable<object[]> BaseTemplates()
        {
            yield return new object[] { "BDR", (short)0 };
            yield return new object[] { "BICenterSite", (short)0 };
            yield return new object[] { "BLANKINTERNET", (short)0 };
            yield return new object[] { "BLOG", (short)0 };
            yield return new object[] { "CMSPUBLISHING", (short)0 };
            yield return new object[] { "COMMUNITY", (short)0 };
            yield return new object[] { "COMMUNITYPORTAL", (short)0 };
            yield return new object[] { "CONTENTCTR", (short)0 };
            yield return new object[] { "DEV", (short)0 };
            yield return new object[] { "EHS", (short)1 };
            yield return new object[] { "ENTERWIKI", (short)0 };
            yield return new object[] { "GROUP", (short)0 };
            yield return new object[] { "OFFILE", (short)1 };
            yield return new object[] { "PROJECTSITE", (short)0 };
            yield return new object[] { "SITEPAGEPUBLISHING", (short)0 };
            yield return new object[] { "SRCHCEN", (short)0 };
            yield return new object[] { "SRCHCENTERLITE", (short)0 };
            yield return new object[] { "STS", (short)0 };
            yield return new object[] { "STS", (short)3 };
            yield return new object[] { "TEAMCHANNEL", (short)0 };
            yield return new object[] { "TEAMCHANNEL", (short)1 };
            yield return new object[] { "VISPRUS", (short)0 };
        }

        [TestMethod]
        [TestCategory("Offline")]
        [DynamicData(nameof(BaseTemplates), DynamicDataSourceType.Method)]
        public void GetBaseTemplate_LoadsAndDeserializesEveryShippedTemplate(string webTemplate, short configuration)
        {
            // A null context is legitimate here - it is only used for logging.
            ProvisioningTemplate template = BaseTemplateManager.GetBaseTemplate(null, webTemplate, configuration);

            Assert.IsNotNull(template,
                $"Base template {webTemplate}{configuration} did not load. Either the embedded resource is " +
                $"missing from the csproj, or its name does not match the expected " +
                $"PnP.Core.Provisioning.BaseTemplates.SPO.{webTemplate}{configuration}Template.xml.");

            // A template that parsed but produced nothing usually means the formatter was resolved
            // for the wrong schema namespace, which does not throw.
            Assert.IsTrue(
                template.Lists.Count > 0 || template.SiteFields.Count > 0 || template.ContentTypes.Count > 0 || template.Features != null,
                $"Base template {webTemplate}{configuration} deserialized to an empty template.");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void GetBaseTemplate_ReturnsNullForAnUnknownSiteTemplate()
        {
            // Not every site template ships a base template, and that must degrade to "extract
            // everything" rather than throwing.
            ProvisioningTemplate template = BaseTemplateManager.GetBaseTemplate(null, "NOSUCHTEMPLATE", 0);

            Assert.IsNull(template);
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void EveryEmbeddedBaseTemplateIsCoveredByThisTest()
        {
            // Keeps the list above honest: if a template is added to the csproj without a row here,
            // this fails rather than the new template going untested.
            const string prefix = "PnP.Core.Provisioning.BaseTemplates.SPO.";

            List<string> embedded = typeof(BaseTemplateManager).Assembly
                .GetManifestResourceNames()
                .Where(n => n.StartsWith(prefix, System.StringComparison.Ordinal))
                .OrderBy(n => n, System.StringComparer.Ordinal)
                .ToList();

            List<string> covered = BaseTemplates()
                .Select(row => $"{prefix}{row[0]}{row[1]}Template.xml")
                .OrderBy(n => n, System.StringComparer.Ordinal)
                .ToList();

            CollectionAssert.AreEqual(embedded, covered,
                $"Embedded base templates and the tested set differ.{System.Environment.NewLine}" +
                $"Embedded: {string.Join(", ", embedded.Except(covered))}{System.Environment.NewLine}" +
                $"Tested only: {string.Join(", ", covered.Except(embedded))}");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void GetConfiguration_SplitsTheWebTemplateConfigurationString()
        {
            // PnP Core carries "STS#0" where CSOM had Web.WebTemplate plus a separate short.
            Assert.AreEqual((short)0, InvokeGetConfiguration("STS#0"));
            Assert.AreEqual((short)3, InvokeGetConfiguration("STS#3"));
            Assert.AreEqual((short)1, InvokeGetConfiguration("EHS#1"));

            // A bare number, which is what some sites report
            Assert.AreEqual((short)0, InvokeGetConfiguration("0"));

            // Nothing usable falls back to 0 rather than throwing - the caller only ever wants a
            // resource name out of this.
            Assert.AreEqual((short)0, InvokeGetConfiguration(null));
            Assert.AreEqual((short)0, InvokeGetConfiguration(string.Empty));
            Assert.AreEqual((short)0, InvokeGetConfiguration("STS#not-a-number"));
        }

        private static short InvokeGetConfiguration(string value)
        {
            MethodInfo method = typeof(BaseTemplateManager).GetMethod("GetConfiguration", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.IsNotNull(method, "BaseTemplateManager.GetConfiguration was renamed or removed.");
            return (short)method.Invoke(null, new object[] { value });
        }
    }
}
