using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Services.Core.CSOM;
using PnP.Core.Provisioning.Services.Core.CSOM.Requests.WebParts;
using PnP.Core.Provisioning.Services.Core.CSOM.Requests.Workflows;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.Test.Live.Csom
{
    /// <summary>
    /// Live round trips for the classic web part (<b>T18</b>) and SharePoint 2013 workflow
    /// (<b>T16</b>) CSOM requests.
    /// </summary>
    /// <remarks>
    /// <para><b>Web parts are tested against a list form page.</b> A modern site's
    /// <c>SitePages/Home.aspx</c> is a client-side page with no classic web part zones, so it would
    /// prove nothing. A list's <c>AllItems.aspx</c> <em>is</em> a classic <c>.aspx</c> with a web
    /// part manager — and it is exactly what <c>ObjectListInstance</c> targets, which makes it the
    /// realistic case rather than a contrived one.</para>
    /// <para><b>Workflows are expected to be unavailable.</b> SP2013 workflow services are retired
    /// in most modern tenants. These report <c>Inconclusive</c> in that case; that is the documented
    /// outcome, not a failure.</para>
    /// </remarks>
    [TestClass]
    public class WebPartAndWorkflowLiveTests : LiveTestBase
    {
        /// <summary>
        /// A minimal, universally available web part - the XML a template would carry.
        /// </summary>
        private const string ContentEditorWebPartXml =
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<webParts>" +
            "<webPart xmlns=\"http://schemas.microsoft.com/WebPart/v3\">" +
            "<metaData>" +
            "<type name=\"Microsoft.SharePoint.WebPartPages.ContentEditorWebPart, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c\" />" +
            "<importErrorMessage>Cannot import this web part.</importErrorMessage>" +
            "</metaData>" +
            "<data>" +
            "<properties>" +
            "<property name=\"Title\" type=\"string\">PnPCoreProvisioningTest WebPart</property>" +
            "<property name=\"ChromeType\" type=\"chrometype\">TitleOnly</property>" +
            "</properties>" +
            "</data>" +
            "</webPart>" +
            "</webParts>";

        #region T18 - classic web parts

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("WebParts")]
        public async Task WebPartLifecycle_AddEnumerateMoveUpdateDelete()
        {
            // Deliberately the CLASSIC site. A modern group-connected site is NoScript, which
            // removes "Add and Customize Pages" and blocks ImportWebPart/AddWebPart - while still
            // allowing GetWebPartDefinitions through the same manager. Running the write path on a
            // modern site tests the site's configuration, not the request.
            using (PnPContext context = await GetClassicContextAsync().ConfigureAwait(false))
            {
                if (await IsNoScriptAsync(context).ConfigureAwait(false))
                {
                    Assert.Inconclusive(
                        "The configured classic test site is a NoScript site, so classic web parts cannot be " +
                        "added to it. Point ClassicSTS0TestSite at a genuine STS#0 site, or disable NoScript " +
                        "(Set-PnPSite -NoScriptSite $false), to verify the web part write path.");
                    return;
                }

                (Guid siteId, Guid webId) = await CsomRequestSender.GetSiteAndWebIdAsync(context).ConfigureAwait(false);

                string listTitle = $"{TestPrefix}WP_{DateTime.UtcNow:HHmmssfff}";
                IList list = null;

                try
                {
                    list = await context.Web.Lists.AddAsync(listTitle, ListTemplateType.GenericList).ConfigureAwait(false);
                    await list.LoadAsync(l => l.RootFolder).ConfigureAwait(false);

                    // The list's default view page: a classic .aspx with a web part manager.
                    string pageUrl = $"{list.RootFolder.ServerRelativeUrl}/AllItems.aspx";
                    Console.WriteLine($"Targeting list form page: {pageUrl}");

                    // --- add -------------------------------------------------------------
                    WebPartDefinitionInfo added;
                    try
                    {
                        added = await CsomRequestSender.SendAsync(context,
                            new AddWebPartRequest(siteId, webId, pageUrl, ContentEditorWebPartXml, "Main", 0))
                            .ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        SkipIfUnavailable("Classic web parts on a list form page", ex);
                        return;
                    }

                    Assert.IsNotNull(added, "AddWebPartRequest returned no result.");
                    Assert.AreNotEqual(Guid.Empty, added.Id, "The added web part has no id.");
                    Console.WriteLine($"  added web part {added.Id} in zone '{added.ZoneId}'");

                    // --- enumerate -------------------------------------------------------
                    List<WebPartDefinitionInfo> definitions = await CsomRequestSender.SendAsync(context,
                        new GetWebPartDefinitionsRequest(siteId, webId, pageUrl)).ConfigureAwait(false);

                    Assert.IsTrue(definitions.Any(d => d.Id == added.Id),
                        $"The added web part {added.Id} was not found when enumerating. Found: " +
                        string.Join(", ", definitions.Select(d => d.Id)));

                    // --- update ----------------------------------------------------------
                    await CsomRequestSender.SendAsync(context,
                        new SaveWebPartPropertiesRequest(siteId, webId, pageUrl, added.Id,
                            title: $"{TestPrefix}Renamed", zoneIndex: 1)).ConfigureAwait(false);

                    // --- move ------------------------------------------------------------
                    await CsomRequestSender.SendAsync(context,
                        new MoveWebPartToRequest(siteId, webId, pageUrl, added.Id, "Main", 2)).ConfigureAwait(false);

                    // --- delete ----------------------------------------------------------
                    await CsomRequestSender.SendAsync(context,
                        new DeleteWebPartRequest(siteId, webId, pageUrl, added.Id)).ConfigureAwait(false);

                    List<WebPartDefinitionInfo> afterDelete = await CsomRequestSender.SendAsync(context,
                        new GetWebPartDefinitionsRequest(siteId, webId, pageUrl)).ConfigureAwait(false);

                    Assert.IsFalse(afterDelete.Any(d => d.Id == added.Id),
                        "The web part was still present after DeleteWebPartRequest.");
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

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("WebParts")]
        public async Task GetWebPartDefinitions_ReadsAnExistingListFormPage()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                (Guid siteId, Guid webId) = await CsomRequestSender.GetSiteAndWebIdAsync(context).ConfigureAwait(false);

                string listTitle = $"{TestPrefix}WPRead_{DateTime.UtcNow:HHmmssfff}";
                IList list = null;

                try
                {
                    list = await context.Web.Lists.AddAsync(listTitle, ListTemplateType.GenericList).ConfigureAwait(false);
                    await list.LoadAsync(l => l.RootFolder).ConfigureAwait(false);

                    string pageUrl = $"{list.RootFolder.ServerRelativeUrl}/AllItems.aspx";

                    List<WebPartDefinitionInfo> definitions;
                    try
                    {
                        definitions = await CsomRequestSender.SendAsync(context,
                            new GetWebPartDefinitionsRequest(siteId, webId, pageUrl)).ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        SkipIfUnavailable("Classic web parts on a list form page", ex);
                        return;
                    }

                    Assert.IsNotNull(definitions);

                    // A list form page ships with its own list view web part, so this should not be
                    // empty - which also proves the response parsing works rather than silently
                    // returning nothing.
                    Console.WriteLine($"Web parts on {pageUrl}: {definitions.Count}");
                    foreach (WebPartDefinitionInfo definition in definitions)
                    {
                        Console.WriteLine($"  {definition.Id} zone='{definition.ZoneId}'");
                    }

                    Assert.IsTrue(definitions.Count > 0,
                        "A list form page should carry at least its own list view web part. An empty result " +
                        "suggests the response was not parsed rather than that the page is empty.");
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

        #endregion

        #region T16 - SP2013 workflows

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Workflows")]
        public async Task GetWorkflowDefinitions_EnumeratesOrReportsWorkflowsRetired()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                (Guid siteId, Guid webId) = await CsomRequestSender.GetSiteAndWebIdAsync(context).ConfigureAwait(false);

                List<WorkflowDefinitionInfo> definitions;
                try
                {
                    definitions = await CsomRequestSender.SendAsync(context,
                        new GetWorkflowDefinitionsRequest(siteId, webId)).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    SkipIfUnavailable("SharePoint 2013 workflow services", ex);
                    return;
                }

                Assert.IsNotNull(definitions);

                Console.WriteLine($"Workflow definitions: {definitions.Count}");
                foreach (WorkflowDefinitionInfo definition in definitions)
                {
                    Console.WriteLine($"  {definition.Id} {definition.DisplayName} published={definition.Published}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Workflows")]
        public async Task GetWorkflowSubscriptions_EnumeratesOrReportsWorkflowsRetired()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                (Guid siteId, Guid webId) = await CsomRequestSender.GetSiteAndWebIdAsync(context).ConfigureAwait(false);

                List<WorkflowSubscriptionInfo> subscriptions;
                try
                {
                    subscriptions = await CsomRequestSender.SendAsync(context,
                        new GetWorkflowSubscriptionsRequest(siteId, webId)).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    SkipIfUnavailable("SharePoint 2013 workflow services", ex);
                    return;
                }

                Assert.IsNotNull(subscriptions);
                Console.WriteLine($"Workflow subscriptions: {subscriptions.Count}");
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Workflows")]
        public async Task SaveAndPublishWorkflowDefinition_RoundTripsOrReportsWorkflowsRetired()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                (Guid siteId, Guid webId) = await CsomRequestSender.GetSiteAndWebIdAsync(context).ConfigureAwait(false);

                // The smallest thing the workflow deployment service will accept as XAML. If
                // workflows are retired this never gets that far.
                var definition = new WorkflowDefinitionInfo
                {
                    DisplayName = $"{TestPrefix}Workflow",
                    Description = "Created by a live test - safe to delete",
                    Xaml = "<Activity xmlns=\"http://schemas.microsoft.com/netfx/2009/xaml/activities\" />",
                    RestrictToType = "Universal",
                };

                WorkflowDefinitionInfo saved = null;
                try
                {
                    saved = await CsomRequestSender.SendAsync(context,
                        new SaveWorkflowDefinitionRequest(siteId, webId, definition)).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    SkipIfUnavailable("SharePoint 2013 workflow services", ex);
                    return;
                }

                try
                {
                    Assert.IsNotNull(saved, "SaveWorkflowDefinitionRequest returned no result.");
                    Assert.AreNotEqual(Guid.Empty, saved.Id, "The saved workflow definition has no id.");
                    Console.WriteLine($"Saved workflow definition {saved.Id}");

                    await CsomRequestSender.SendAsync(context,
                        new PublishWorkflowDefinitionRequest(siteId, webId, saved.Id)).ConfigureAwait(false);

                    List<WorkflowDefinitionInfo> definitions = await CsomRequestSender.SendAsync(context,
                        new GetWorkflowDefinitionsRequest(siteId, webId)).ConfigureAwait(false);

                    Assert.IsTrue(definitions.Any(d => d.Id == saved.Id),
                        "The saved workflow definition was not found when enumerating.");
                }
                finally
                {
                    if (saved != null && saved.Id != Guid.Empty)
                    {
                        try
                        {
                            await CsomRequestSender.SendAsync(context,
                                new DeleteWorkflowDefinitionRequest(siteId, webId, saved.Id)).ConfigureAwait(false);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Could not delete workflow definition {saved.Id}: {ex.Message}");
                        }
                    }
                }
            }
        }

        #endregion
    }
}
