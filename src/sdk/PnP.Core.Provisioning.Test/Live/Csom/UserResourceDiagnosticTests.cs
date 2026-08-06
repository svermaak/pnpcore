using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Services.Core.CSOM;
using PnP.Core.Provisioning.Services.Core.CSOM.Requests.UserResources;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.Test.Live.Csom
{
    /// <summary>
    /// Triage for the <b>T6</b> user resource round trip, which failed on first contact.
    /// </summary>
    /// <remarks>
    /// <para><b>The symptom:</b> <c>SetValueForUICulture</c> returns without error, and
    /// <c>GetValueForUICulture</c> then returns the object's <em>default</em> title rather than the
    /// value just written. No exception anywhere.</para>
    /// <para><b>Two candidate causes,</b> and guessing between them would be a mistake:</para>
    /// <list type="number">
    /// <item><description>The CSOM request is malformed and the write never happens.</description></item>
    /// <item><description>The write is correct but SharePoint discards it, because a user resource
    /// only persists for a language in the web's <c>SupportedUILanguageIds</c>. Writing to a
    /// language the site does not have enabled is accepted and dropped.</description></item>
    /// </list>
    /// <para>This test distinguishes them: it reports what languages the site actually supports and
    /// then round-trips against one of <em>those</em>. If that works, the request is fine and the
    /// original test simply picked a language the site does not have - which also means <b>T2
    /// (supported UI languages) is a hard prerequisite for T6</b>, exactly as the backlog says.</para>
    /// </remarks>
    [TestClass]
    public class UserResourceDiagnosticTests : LiveTestBase
    {
        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("UserResources")]
        public async Task Diagnose_WhichLanguagesDoesTheSiteActuallySupport()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                IWeb web = await context.Web.GetAsync(
                    w => w.Language, w => w.IsMultilingual, w => w.SupportedUILanguageIds).ConfigureAwait(false);

                Console.WriteLine($"Web default language LCID : {web.Language} ({SafeCultureName(web.Language)})");
                Console.WriteLine($"IsMultilingual            : {web.IsMultilingual}");
                Console.WriteLine($"SupportedUILanguageIds    : " +
                    (web.SupportedUILanguageIds == null || web.SupportedUILanguageIds.Count == 0
                        ? "<none>"
                        : string.Join(", ", web.SupportedUILanguageIds.Select(l => $"{l} ({SafeCultureName(l)})"))));

                Assert.IsTrue(true, "Diagnostic only - see output.");
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("UserResources")]
        public async Task RoundTrip_AgainstALanguageTheSiteActuallySupports()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                IWeb web = await context.Web.GetAsync(
                    w => w.Language, w => w.SupportedUILanguageIds).ConfigureAwait(false);

                // Prefer a supported language that is NOT the default, since writing a resource for
                // the default language is indistinguishable from doing nothing.
                int? targetLcid = web.SupportedUILanguageIds?
                    .Where(l => l != web.Language)
                    .Cast<int?>()
                    .FirstOrDefault();

                if (targetLcid == null)
                {
                    Assert.Inconclusive(
                        $"This site has no supported UI language other than its default ({web.Language}), so a " +
                        $"localized value cannot be stored and T6 cannot be verified here. This is a site " +
                        $"configuration matter, not a defect: enable a second language on the site (Site Settings > " +
                        $"Language Settings), or run against a multilingual site. It also confirms that backlog T2 " +
                        $"(supported UI languages) is a genuine prerequisite for T6.");
                    return;
                }

                string cultureName = SafeCultureName(targetLcid.Value);
                Console.WriteLine($"Round-tripping against supported language {targetLcid} ({cultureName})");

                (Guid siteId, Guid webId) = await CsomRequestSender.GetSiteAndWebIdAsync(context).ConfigureAwait(false);

                string listTitle = $"{TestPrefix}Diag_{DateTime.UtcNow:HHmmssfff}";
                IList list = null;

                try
                {
                    list = await context.Web.Lists.AddAsync(listTitle, ListTemplateType.GenericList).ConfigureAwait(false);
                    var resource = UserResourcePath.ForList(siteId, webId, list.Id, ResourceProperty.Title);

                    string localized = $"{TestPrefix}Localized_{DateTime.UtcNow:HHmmss}";

                    await CsomRequestSender.SendAsync(context,
                        new SetValueForUICultureRequest(resource, cultureName, localized)).ConfigureAwait(false);

                    string readBack = await CsomRequestSender.SendAsync(context,
                        new GetValueForUICultureRequest(resource, cultureName)).ConfigureAwait(false);

                    Console.WriteLine($"  wrote     : {localized}");
                    Console.WriteLine($"  read back : {readBack}");

                    Assert.AreEqual(localized, readBack,
                        $"Even against a supported language ({cultureName}) the value did not round-trip. " +
                        $"That points at the CSOM request itself rather than site configuration.");
                }
                finally
                {
                    if (list != null)
                    {
                        try
                        {
                            await list.DeleteAsync().ConfigureAwait(false);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Could not delete test list {listTitle}: {ex.Message}");
                        }
                    }
                }
            }
        }

        private static string SafeCultureName(int lcid)
        {
            try
            {
                return new CultureInfo(lcid).Name;
            }
            catch (CultureNotFoundException)
            {
                return "<unknown>";
            }
        }
    }
}
