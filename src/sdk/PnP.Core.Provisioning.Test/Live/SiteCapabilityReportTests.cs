using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.ObjectHandlers.Utilities;
using PnP.Core.Provisioning.Test.Utilities;
using PnP.Core.Services;
using System;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.Test.Live
{
    /// <summary>
    /// Reports what each configured site actually is, so verification gaps can be attributed to the
    /// right cause.
    /// </summary>
    /// <remarks>
    /// <para>Several handlers behave differently on NoScript sites, on modern versus classic sites,
    /// and on multilingual ones. When a test skips, the reason is usually one of those - and
    /// guessing which is how a real defect gets filed as "tenant limitation".</para>
    /// <para>It also catches a subtler trap: <c>TestCommonBase</c> silently falls back to
    /// <c>TestSite</c> when another site's url is not configured. A test that thinks it is running
    /// against the classic site may not be, and would then draw conclusions about the wrong site.
    /// This prints the resolved url so that cannot go unnoticed.</para>
    /// </remarks>
    [TestClass]
    public class SiteCapabilityReportTests : LiveTestBase
    {
        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Report")]
        public async Task ReportConfiguredSiteCapabilities()
        {
            await ReportAsync(TestCommon.TestSite, "TestSite").ConfigureAwait(false);
            await ReportAsync(TestCommon.ClassicSTS0TestSite, "ClassicSTS0TestSite").ConfigureAwait(false);
            await ReportAsync(TestCommon.NoGroupTestSite, "NoGroupTestSite").ConfigureAwait(false);
        }

        private static async Task ReportAsync(string configurationName, string label)
        {
            Console.WriteLine();
            Console.WriteLine($"=== {label} ===");

            try
            {
                using (PnPContext context = await TestCommon.Instance.GetContextAsync(configurationName).ConfigureAwait(false))
                {
                    IWeb web = await context.Web.GetAsync(
                        w => w.Url, w => w.Title, w => w.WebTemplate, w => w.WebTemplateConfiguration,
                        w => w.Language, w => w.IsMultilingual).ConfigureAwait(false);

                    bool noScript = await context.Web.IsNoScriptSiteAsync().ConfigureAwait(false);
                    bool isComm = await SiteTypeHelper.IsCommunicationSiteAsync(context).ConfigureAwait(false);
                    bool isTeam = await SiteTypeHelper.IsModernTeamSiteAsync(context).ConfigureAwait(false);

                    Console.WriteLine($"  Url          : {web.Url}");
                    Console.WriteLine($"  Template     : {web.WebTemplateConfiguration} ({web.WebTemplate})");
                    Console.WriteLine($"  Kind         : {(isComm ? "communication" : isTeam ? "modern team" : "classic")}");
                    Console.WriteLine($"  NoScript     : {noScript}   <-- blocks property bag / web part / custom action WRITES");
                    Console.WriteLine($"  Multilingual : {web.IsMultilingual} (default LCID {web.Language})");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  COULD NOT REACH: {Describe(ex)}");
            }
        }
    }
}
