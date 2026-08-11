using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using CanvasControlModel = PnP.Core.Provisioning.Model.CanvasControl;
using CanvasSectionModel = PnP.Core.Provisioning.Model.CanvasSection;
using ClientSidePageModel = PnP.Core.Provisioning.Model.ClientSidePage;
using FooterModel = PnP.Core.Provisioning.Model.SiteFooter;
using HeaderModel = PnP.Core.Provisioning.Model.SiteHeader;
using ThemeModel = PnP.Core.Provisioning.Model.Theme;

namespace PnP.Core.Provisioning.Test.Scenarios
{
    /// <summary>
    /// Scenario 2 - modern client-side pages, site chrome and a theme.
    /// </summary>
    /// <remarks>
    /// <para>The test strategy calls this the area where PnP Core is strongest and expects it to be
    /// the easiest green. It is worth saying why that is not the same as "low value": these four
    /// handlers are the ones a modern site actually uses, and they are the part of the engine most
    /// likely to be exercised by real callers on day one.</para>
    /// <para><b>The page assertions read the page back, not the template.</b> A client-side page's
    /// content lives in a canvas the engine writes as JSON; a page that exists with an empty canvas
    /// is the failure mode, and it is invisible unless the sections and controls are counted.</para>
    /// </remarks>
    [TestClass]
    public class Scenario02PagesAndChromeTests : ScenarioTestBase
    {
        private const string Prefix = "PnPCoreScenario2_";

        private static readonly string PageName = $"{Prefix}Home.aspx";

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Scenario")]
        [Timeout(45 * 60 * 1000)]
        public async Task Scenario2_PagesHeaderFooterAndTheme()
        {
            await RunScenarioAsync("s2", BuildTemplate(), new[]
            {
                ConfigurationHandler.Pages,
                ConfigurationHandler.SiteHeader,
                ConfigurationHandler.SiteFooter,
            },
            AssertAsync,

            // Asked for explicitly, like scenario 5's term groups. Extracting a site reports its
            // home page and nothing else by default, which is right - a template carrying every page
            // on a site is rarely what someone extracting one wants.
            configuration => configuration.Pages.IncludeAllClientSidePages = true)
            .ConfigureAwait(false);
        }

        private static ProvisioningTemplate BuildTemplate()
        {
            var template = new ProvisioningTemplate { Id = "SCENARIO-2" };

            var page = new ClientSidePageModel
            {
                PageName = PageName,
                Title = $"{Prefix}Home",
                Layout = nameof(PageLayoutType.Article),
                Overwrite = true,
                Publish = true,
                EnableComments = false,
            };

            // Two columns with a control in each: a single control would not distinguish "the canvas
            // was written" from "a control landed in the default section".
            var section = new CanvasSectionModel { Type = CanvasSectionType.TwoColumn, Order = 1 };

            section.Controls.Add(new CanvasControlModel
            {
                Type = WebPartType.Text,
                Column = 1,
                Order = 1,
                ControlProperties = new Dictionary<string, string> { { "Text", "<p>Scenario 2 left</p>" } },
            });

            section.Controls.Add(new CanvasControlModel
            {
                Type = WebPartType.Text,
                Column = 2,
                Order = 1,
                ControlProperties = new Dictionary<string, string> { { "Text", "<p>Scenario 2 right</p>" } },
            });

            page.Sections.Add(section);
            template.ClientSidePages.Add(page);

            template.Header = new HeaderModel
            {
                Layout = SiteHeaderLayout.Compact,
                MenuStyle = SiteHeaderMenuStyle.MegaMenu,
                BackgroundEmphasis = Emphasis.Strong,
            };

            // The footer is why the harness creates a communication site: a group connected site has
            // none, and the handler correctly does nothing there - which would make a green run here
            // prove nothing at all.
            template.Footer = new FooterModel
            {
                Enabled = true,
                Layout = SiteFooterLayout.Extended,
                BackgroundEmphasis = Emphasis.Neutral,
                RemoveExistingNodes = false,
            };

            template.Theme = new ThemeModel { Name = nameof(SharePointTheme.Blue) };

            return template;
        }

        private static async Task AssertAsync(ProvisioningTemplate extracted, PnPContext site)
        {
            // ------------------------------------------------------------------------------------
            // The page, read back rather than trusted.
            // ------------------------------------------------------------------------------------
            IPage page = (await site.Web.GetPagesAsync(PageName).ConfigureAwait(false)).FirstOrDefault();

            Assert.IsNotNull(page, $"The page '{PageName}' was not created.");

            Console.WriteLine($"Page has {page.Sections.Count} section(s) and {page.Controls.Count} control(s)");

            Assert.AreEqual(1, page.Sections.Count, "The page's canvas has the wrong number of sections.");

            // The count is the assertion that means something: a page created with an empty canvas
            // looks identical to one created correctly until it is opened.
            Assert.AreEqual(2, page.Controls.Count,
                "The page was created but its controls were not written - an empty canvas is the " +
                "failure mode here, and it is invisible from the page's existence alone.");

            // ------------------------------------------------------------------------------------
            // The chrome.
            // ------------------------------------------------------------------------------------
            await site.Web.LoadAsync(w => w.HeaderLayout, w => w.MegaMenuEnabled, w => w.FooterEnabled)
                .ConfigureAwait(false);

            Console.WriteLine($"Header layout: {site.Web.HeaderLayout}, mega menu: {site.Web.MegaMenuEnabled}, " +
                $"footer: {site.Web.FooterEnabled}");

            Assert.AreEqual(HeaderLayoutType.Compact, site.Web.HeaderLayout, "The header layout was not applied.");
            Assert.IsTrue(site.Web.MegaMenuEnabled, "The mega menu was not enabled.");
            Assert.IsTrue(site.Web.FooterEnabled, "The footer was not enabled.");

            // ------------------------------------------------------------------------------------
            // And the extract found the page.
            // ------------------------------------------------------------------------------------
            ClientSidePageModel extractedPage = extracted.ClientSidePages
                .FirstOrDefault(p => string.Equals(p.PageName, PageName, StringComparison.OrdinalIgnoreCase));

            Assert.IsNotNull(extractedPage,
                $"The extract did not report the page. Found: " +
                string.Join(", ", extracted.ClientSidePages.Select(p => p.PageName)));

            // Controls, not sections: an extract that produced sections and no controls would satisfy
            // a section count and still have lost the page's content.
            Assert.IsTrue(extractedPage.Sections.Sum(s => s.Controls.Count) >= 2,
                "The extract reported the page but not its controls, so its canvas was not read back. " +
                $"Sections: {extractedPage.Sections.Count}, controls: " +
                $"{extractedPage.Sections.Sum(s => s.Controls.Count)}");
        }
    }
}
