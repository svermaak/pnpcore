using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.ObjectHandlers.Utilities;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using PageLayoutModel = PnP.Core.Provisioning.Model.PageLayout;

namespace PnP.Core.Provisioning.Test.Offline.ObjectHandlers
{
    /// <summary>
    /// Pins the publishing property-bag XML that <c>ObjectPublishing</c> writes and reads.
    /// </summary>
    /// <remarks>
    /// <para><b>Why offline.</b> The three settings behind <c>&lt;pnp:Publishing&gt;</c> have no API -
    /// they are XML strings in the web's property bag that the SharePoint UI parses. So the shape is
    /// a contract with SharePoint, not a format this engine chooses, and every value below is checked
    /// against what PnP Framework emits in
    /// <c>PnPF:Extensions/BrandingExtensions.cs</c> (<c>SetAvailableWebTemplates</c>,
    /// <c>SetAvailablePageLayouts</c>, <c>CreateXmlNodeFromPageLayout</c>) and reads back in
    /// <c>PnPF:Provisioning/ObjectHandlers/ObjectPublishing.cs</c> (<c>GetAvailablePageLayouts</c>).</para>
    /// <para><b>Why it earns its place.</b> Reaching the write path live needs a site with the
    /// publishing feature active. Even with one, a live apply-then-extract would pass with the wrong
    /// element names as long as the writer and the reader agreed with each other - and it would pass
    /// with the gallery prefix left on, because both sides would carry it. Neither bug is visible to a
    /// round-trip against itself; both are visible here.</para>
    /// </remarks>
    [TestClass]
    [TestCategory("Offline")]
    public class PublishingPropertyBagXmlTests
    {
        private static string Passthrough(string value) => value;

        #region Web templates

        [TestMethod]
        public void BuildWebTemplates_MatchesTheShapeSharePointReads()
        {
            var templates = new List<AvailableWebTemplate>
            {
                new AvailableWebTemplate { LanguageCode = 1033, TemplateName = "STS#0" },
                new AvailableWebTemplate { LanguageCode = 1033, TemplateName = "BLOG#0" },
            };

            string xml = PublishingPropertyBagXml.BuildWebTemplates(templates, Passthrough);

            Assert.AreEqual(
                "<webtemplates><lcid id=\"1033\"><webtemplate name=\"STS#0\" /><webtemplate name=\"BLOG#0\" /></lcid></webtemplates>",
                xml);
        }

        [TestMethod]
        public void BuildWebTemplates_GroupsByLanguage()
        {
            // The same template offered in two languages has to land under two <lcid> elements, not
            // two <webtemplates> roots - SharePoint reads the nesting, not the order.
            var templates = new List<AvailableWebTemplate>
            {
                new AvailableWebTemplate { LanguageCode = 1033, TemplateName = "STS#0" },
                new AvailableWebTemplate { LanguageCode = 1043, TemplateName = "STS#0" },
                new AvailableWebTemplate { LanguageCode = 1033, TemplateName = "BLOG#0" },
            };

            XElement root = XElement.Parse(PublishingPropertyBagXml.BuildWebTemplates(templates, Passthrough));

            Assert.AreEqual(2, root.Elements("lcid").Count());
            Assert.AreEqual(2, root.Elements("lcid").First(e => (string)e.Attribute("id") == "1033")
                .Elements("webtemplate").Count());
            Assert.AreEqual(1, root.Elements("lcid").First(e => (string)e.Attribute("id") == "1043")
                .Elements("webtemplate").Count());
        }

        [TestMethod]
        public void BuildWebTemplates_ResolvesTokensInTheTemplateName()
        {
            var templates = new List<AvailableWebTemplate>
            {
                new AvailableWebTemplate { LanguageCode = 1033, TemplateName = "{parameter:tpl}" },
            };

            string xml = PublishingPropertyBagXml.BuildWebTemplates(
                templates, v => v == "{parameter:tpl}" ? "STS#0" : v);

            StringAssert.Contains(xml, "name=\"STS#0\"");
        }

        [TestMethod]
        public void ReadWebTemplates_RoundTripsWhatWasBuilt()
        {
            var templates = new List<AvailableWebTemplate>
            {
                new AvailableWebTemplate { LanguageCode = 1033, TemplateName = "STS#0" },
                new AvailableWebTemplate { LanguageCode = 1043, TemplateName = "BLOG#0" },
            };

            List<AvailableWebTemplate> read = PublishingPropertyBagXml.ReadWebTemplates(
                PublishingPropertyBagXml.BuildWebTemplates(templates, Passthrough));

            CollectionAssert.AreEquivalent(
                templates.Select(t => $"{t.LanguageCode}:{t.TemplateName}").ToList(),
                read.Select(t => $"{t.LanguageCode}:{t.TemplateName}").ToList());
        }

        [TestMethod]
        public void ReadWebTemplates_IsEmptyForAnUnsetProperty()
        {
            // "Allow all" is stored as an empty string, not as an absent key - see
            // ClearAvailableWebTemplates in PnP Framework.
            Assert.AreEqual(0, PublishingPropertyBagXml.ReadWebTemplates(null).Count);
            Assert.AreEqual(0, PublishingPropertyBagXml.ReadWebTemplates(string.Empty).Count);
        }

        [TestMethod]
        public void ReadWebTemplates_SkipsANonNumericLanguage()
        {
            // PnP Framework keys the group "all" when a template carries no language code. Nothing in
            // the PnP Core model can express that, so such a group is dropped rather than turned into
            // LanguageCode 0, which would be a real language on re-apply.
            List<AvailableWebTemplate> read = PublishingPropertyBagXml.ReadWebTemplates(
                "<webtemplates><lcid id=\"all\"><webtemplate name=\"STS#0\" /></lcid>" +
                "<lcid id=\"1033\"><webtemplate name=\"BLOG#0\" /></lcid></webtemplates>");

            Assert.AreEqual(1, read.Count);
            Assert.AreEqual("BLOG#0", read[0].TemplateName);
            Assert.AreEqual(1033, read[0].LanguageCode);
        }

        #endregion

        #region Page layouts

        [TestMethod]
        public void BuildPageLayouts_MatchesTheShapeSharePointReads()
        {
            // The comment on CreateXmlNodeFromPageLayout in PnP Framework spells this out verbatim:
            // <layout guid="944ea6be-f287-42c6-aa11-3fd75ab1ee9e" url="_catalogs/masterpage/ArticleLeft.aspx" />
            string xml = PublishingPropertyBagXml.BuildPageLayouts(new[]
            {
                PublishingPropertyBagXml.BuildLayout(
                    "944ea6be-f287-42c6-aa11-3fd75ab1ee9e", "_catalogs/masterpage/ArticleLeft.aspx"),
            });

            Assert.AreEqual(
                "<pagelayouts><layout guid=\"944ea6be-f287-42c6-aa11-3fd75ab1ee9e\" " +
                "url=\"_catalogs/masterpage/ArticleLeft.aspx\" /></pagelayouts>",
                xml);
        }

        [TestMethod]
        public void ReadPageLayouts_StripsTheGalleryPrefix()
        {
            // The property bag stores a site-relative url; a template's Path is the bare name. Leaving
            // the prefix on would make every extracted template differ from PnP Framework's.
            List<PageLayoutModel> layouts = PublishingPropertyBagXml.ReadPageLayouts(
                "<pagelayouts><layout guid=\"g1\" url=\"_catalogs/masterpage/ArticleLeft.aspx\" /></pagelayouts>",
                null);

            Assert.AreEqual(1, layouts.Count);
            Assert.AreEqual("ArticleLeft.aspx", layouts[0].Path);
        }

        [TestMethod]
        public void ReadPageLayouts_MarksTheDefault()
        {
            // This is the assertion the stripping bug broke: the default url is stripped, so an
            // unstripped path can never equal it and IsDefault is false for every layout.
            List<PageLayoutModel> layouts = PublishingPropertyBagXml.ReadPageLayouts(
                "<pagelayouts>" +
                "<layout guid=\"g1\" url=\"_catalogs/masterpage/ArticleLeft.aspx\" />" +
                "<layout guid=\"g2\" url=\"_catalogs/masterpage/ArticleRight.aspx\" />" +
                "</pagelayouts>",
                "<layout guid=\"g2\" url=\"_catalogs/masterpage/ArticleRight.aspx\" />");

            Assert.AreEqual(2, layouts.Count);
            Assert.IsFalse(layouts.Single(l => l.Path == "ArticleLeft.aspx").IsDefault);
            Assert.IsTrue(layouts.Single(l => l.Path == "ArticleRight.aspx").IsDefault);
        }

        [TestMethod]
        public void ReadPageLayouts_HasNoDefaultWhenTheWebInherits()
        {
            // __inherit is a literal, not XML. Parsing it throws, and the handler's catch is at debug
            // level - so getting this wrong loses the layouts silently.
            List<PageLayoutModel> layouts = PublishingPropertyBagXml.ReadPageLayouts(
                "<pagelayouts><layout guid=\"g1\" url=\"_catalogs/masterpage/ArticleLeft.aspx\" /></pagelayouts>",
                PublishingPropertyBagXml.InheritMarker);

            Assert.AreEqual(1, layouts.Count);
            Assert.IsFalse(layouts[0].IsDefault);
        }

        [TestMethod]
        public void ReadPageLayouts_IsEmptyWhenTheListItselfInherits()
        {
            Assert.AreEqual(0, PublishingPropertyBagXml.ReadPageLayouts(
                PublishingPropertyBagXml.InheritMarker, null).Count);
        }

        [TestMethod]
        public void ReadPageLayouts_IsEmptyWhenAllLayoutsAreAllowed()
        {
            // AllowAllPageLayouts writes an empty string.
            Assert.AreEqual(0, PublishingPropertyBagXml.ReadPageLayouts(string.Empty, null).Count);
            Assert.AreEqual(0, PublishingPropertyBagXml.ReadPageLayouts(null, null).Count);
        }

        [TestMethod]
        public void ReadPageLayouts_SkipsALayoutWithNoUrl()
        {
            List<PageLayoutModel> layouts = PublishingPropertyBagXml.ReadPageLayouts(
                "<pagelayouts><layout guid=\"g1\" /><layout guid=\"g2\" url=\"Article.aspx\" /></pagelayouts>",
                null);

            Assert.AreEqual(1, layouts.Count);
            Assert.AreEqual("Article.aspx", layouts[0].Path);
        }

        [TestMethod]
        public void ReadDefaultPageLayoutUrl_SurvivesAMalformedValue()
        {
            Assert.IsNull(PublishingPropertyBagXml.ReadDefaultPageLayoutUrl("<layout guid=\"g1\""));
            Assert.IsNull(PublishingPropertyBagXml.ReadDefaultPageLayoutUrl("<layout guid=\"g1\" />"));
        }

        [TestMethod]
        public void PageLayouts_RoundTripBuildThenRead()
        {
            XElement left = PublishingPropertyBagXml.BuildLayout("g1", "_catalogs/masterpage/ArticleLeft.aspx");
            XElement right = PublishingPropertyBagXml.BuildLayout("g2", "_catalogs/masterpage/ArticleRight.aspx");

            List<PageLayoutModel> layouts = PublishingPropertyBagXml.ReadPageLayouts(
                PublishingPropertyBagXml.BuildPageLayouts(new[] { left, right }),
                right.ToString(SaveOptions.DisableFormatting));

            CollectionAssert.AreEqual(
                new[] { "ArticleLeft.aspx", "ArticleRight.aspx" },
                layouts.Select(l => l.Path).ToArray());
            Assert.IsTrue(layouts[1].IsDefault);
        }

        #endregion

        #region Keys

        [TestMethod]
        public void TheKeysAreTheOnesSharePointReads()
        {
            // Spelled out because a typo here fails silently: the write succeeds, the UI ignores it.
            Assert.AreEqual("__WebTemplates", PublishingPropertyBagXml.AvailableWebTemplatesKey);
            Assert.AreEqual("__InheritWebTemplates", PublishingPropertyBagXml.InheritWebTemplatesKey);
            Assert.AreEqual("__PageLayouts", PublishingPropertyBagXml.AvailablePageLayoutsKey);
            Assert.AreEqual("__DefaultPageLayout", PublishingPropertyBagXml.DefaultPageLayoutKey);
            Assert.AreEqual("__inherit", PublishingPropertyBagXml.InheritMarker);
        }

        #endregion
    }
}
