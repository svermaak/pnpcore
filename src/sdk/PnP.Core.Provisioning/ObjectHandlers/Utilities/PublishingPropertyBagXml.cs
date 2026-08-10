using PnP.Core.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using PageLayoutModel = PnP.Core.Provisioning.Model.PageLayout;

namespace PnP.Core.Provisioning.ObjectHandlers.Utilities
{
    /// <summary>
    /// Builds and reads the property bag XML SharePoint uses for publishing settings.
    /// </summary>
    /// <remarks>
    /// <para><b>These three settings have no API.</b> <c>__WebTemplates</c>, <c>__PageLayouts</c>
    /// and <c>__DefaultPageLayout</c> are XML strings in the web's property bag, and PnP Framework's
    /// "extension methods" for them are just builders for that XML. The SharePoint UI is what reads
    /// it back, so the shape is a contract - not a format this engine may choose.</para>
    /// <para><b>Separated from the handler on purpose.</b> Reaching the write path needs a site with
    /// the publishing feature active, which the test tenant does not have; keeping the format here,
    /// free of any <c>PnPContext</c>, means the part most likely to be wrong is covered by offline
    /// tests that run anywhere. What is left in the handler is the property-bag write itself, which
    /// is a one-liner that half a dozen other handlers already exercise.</para>
    /// </remarks>
    internal static class PublishingPropertyBagXml
    {
        internal const string AvailableWebTemplatesKey = "__WebTemplates";
        internal const string InheritWebTemplatesKey = "__InheritWebTemplates";
        internal const string AvailablePageLayoutsKey = "__PageLayouts";
        internal const string DefaultPageLayoutKey = "__DefaultPageLayout";

        /// <summary>
        /// The value <c>__DefaultPageLayout</c> or <c>__PageLayouts</c> carries when the web inherits
        /// from its parent rather than declaring its own.
        /// </summary>
        internal const string InheritMarker = "__inherit";

        /// <summary>
        /// The gallery prefix stripped off an extracted layout path.
        /// </summary>
        /// <remarks>
        /// The property bag stores a site-relative url - <c>_catalogs/masterpage/ArticleLeft.aspx</c> -
        /// but a template's <c>Path</c> carries the bare name, so PnP Framework strips this on the way
        /// out. Keeping it would not only make extracted templates differ from PnP Framework's, it
        /// would break <c>IsDefault</c>: the default url is stripped, so an unstripped path could
        /// never equal it.
        /// </remarks>
        private const string GalleryPrefix = "_catalogs/masterpage/";

        #region Web templates

        /// <summary>
        /// Builds <c>__WebTemplates</c>: the templates a subsite may be created from.
        /// </summary>
        /// <remarks>
        /// Grouped by LCID because SharePoint's format nests <c>&lt;webtemplate&gt;</c> under
        /// <c>&lt;lcid&gt;</c> - the same template name can be offered in several languages.
        /// </remarks>
        internal static string BuildWebTemplates(IEnumerable<AvailableWebTemplate> templates, Func<string, string> parse)
        {
            var root = new XElement("webtemplates");

            foreach (IGrouping<int, AvailableWebTemplate> byLanguage in templates.GroupBy(t => t.LanguageCode))
            {
                var lcid = new XElement("lcid",
                    new XAttribute("id", byLanguage.Key.ToString(CultureInfo.InvariantCulture)));

                foreach (AvailableWebTemplate template in byLanguage)
                {
                    lcid.Add(new XElement("webtemplate",
                        new XAttribute("name", parse(template.TemplateName))));
                }

                root.Add(lcid);
            }

            return root.ToString(SaveOptions.DisableFormatting);
        }

        /// <summary>
        /// Reads <c>__WebTemplates</c> back.
        /// </summary>
        internal static List<AvailableWebTemplate> ReadWebTemplates(string xml)
        {
            var templates = new List<AvailableWebTemplate>();

            if (string.IsNullOrWhiteSpace(xml))
            {
                return templates;
            }

            foreach (XElement lcid in XElement.Parse(xml).Elements("lcid"))
            {
                if (!int.TryParse((string)lcid.Attribute("id"), NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out int languageCode))
                {
                    continue;
                }

                foreach (XElement webTemplate in lcid.Elements("webtemplate"))
                {
                    templates.Add(new AvailableWebTemplate
                    {
                        LanguageCode = languageCode,
                        TemplateName = (string)webTemplate.Attribute("name"),
                    });
                }
            }

            return templates;
        }

        #endregion

        #region Page layouts

        /// <summary>
        /// One <c>&lt;layout&gt;</c> element - a page layout's unique id and its site-relative url.
        /// </summary>
        /// <remarks>
        /// The id is what SharePoint matches on, not the url, which is why the catalog has to be
        /// consulted rather than the template's path used directly.
        /// </remarks>
        internal static XElement BuildLayout(string uniqueId, string siteRelativeUrl)
        {
            return new XElement("layout",
                new XAttribute("guid", uniqueId),
                new XAttribute("url", siteRelativeUrl));
        }

        /// <summary>
        /// Wraps layout elements into <c>__PageLayouts</c>.
        /// </summary>
        internal static string BuildPageLayouts(IEnumerable<XElement> layouts)
        {
            var root = new XElement("pagelayouts");

            foreach (XElement layout in layouts)
            {
                root.Add(layout);
            }

            return root.ToString(SaveOptions.DisableFormatting);
        }

        /// <summary>
        /// Reads <c>__PageLayouts</c> and <c>__DefaultPageLayout</c> back into the template's shape.
        /// </summary>
        /// <param name="availableXml">The <c>__PageLayouts</c> value</param>
        /// <param name="defaultXml">The <c>__DefaultPageLayout</c> value, or the inherit marker</param>
        internal static List<PageLayoutModel> ReadPageLayouts(string availableXml, string defaultXml)
        {
            var layouts = new List<PageLayoutModel>();

            // Empty means "allow all", and the inherit marker means "ask the parent". Neither is a
            // list of layouts, and neither is XML - parsing the marker would throw.
            if (string.IsNullOrWhiteSpace(availableXml)
                || availableXml.Trim().Equals(InheritMarker, StringComparison.OrdinalIgnoreCase))
            {
                return layouts;
            }

            string defaultUrl = ReadDefaultPageLayoutUrl(defaultXml);

            foreach (XElement layout in XElement.Parse(availableXml).Descendants("layout"))
            {
                var url = (string)layout.Attribute("url");

                if (url == null)
                {
                    continue;
                }

                string path = StripGalleryPrefix(url);

                layouts.Add(new PageLayoutModel
                {
                    Path = path,
                    IsDefault = defaultUrl != null
                        && string.Equals(path, defaultUrl, StringComparison.OrdinalIgnoreCase),
                });
            }

            return layouts;
        }

        /// <summary>
        /// The layout name in <c>__DefaultPageLayout</c>, or null when the web inherits or has none.
        /// </summary>
        /// <remarks>
        /// The value is either a <c>&lt;layout&gt;</c> element or the literal <c>__inherit</c> -
        /// parsing the latter as XML throws, which is why it is checked first. The gallery prefix is
        /// stripped so the result is directly comparable to a layout's extracted path.
        /// </remarks>
        internal static string ReadDefaultPageLayoutUrl(string defaultXml)
        {
            if (string.IsNullOrWhiteSpace(defaultXml)
                || defaultXml.Trim().Equals(InheritMarker, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            try
            {
                var url = (string)XElement.Parse(defaultXml).Attribute("url");
                return url == null ? null : StripGalleryPrefix(url);
            }
            catch (System.Xml.XmlException)
            {
                return null;
            }
        }

        private static string StripGalleryPrefix(string url)
        {
            int index = url.IndexOf(GalleryPrefix, StringComparison.OrdinalIgnoreCase);
            return index < 0 ? url : url.Remove(index, GalleryPrefix.Length);
        }

        #endregion
    }
}
