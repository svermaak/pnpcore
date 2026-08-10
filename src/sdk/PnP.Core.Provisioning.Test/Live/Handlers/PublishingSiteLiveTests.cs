using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Admin.Model.SharePoint;
using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.Provisioning.ObjectHandlers;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using PageLayoutModel = PnP.Core.Provisioning.Model.PageLayout;
using PublishingModel = PnP.Core.Provisioning.Model.Publishing;
using TimeZone = PnP.Core.Admin.Model.SharePoint.TimeZone;

namespace PnP.Core.Provisioning.Test.Live.Handlers
{
    /// <summary>
    /// The one test that reaches <c>ObjectPublishing</c>'s write path (<b>T17</b>), on a site
    /// created for the purpose.
    /// </summary>
    /// <remarks>
    /// <para><b>Why this test exists.</b> Everything else about publishing was verified against the
    /// existing test sites - the NoScript guard, the feature gate, the refusal when publishing is
    /// off. What could not be reached is the part that only runs when publishing <em>is</em> on:
    /// activating the two features in the right order, reading the master page gallery, and writing
    /// the four property bag keys. No site in the test tenant has publishing, and turning it on for
    /// an existing one is not cleanly reversible - so this test brings its own site and deletes it.</para>
    /// <para><b>It is deliberately one method.</b> Creating and deleting a site collection costs
    /// minutes, and the assertions are a sequence - publishing must be activated before the gallery
    /// can be read, and the gallery must be read before the layouts are known. Splitting them across
    /// methods would mean either several sites or shared mutable state between tests.</para>
    /// <para><b>The XML shape itself is not asserted here.</b> That is
    /// <c>PublishingPropertyBagXmlTests</c>, offline - a live round trip through the engine's own
    /// writer and reader agrees with itself no matter what shape it picks. What this test adds is
    /// that SharePoint accepts the write, that the gallery lookup finds real layouts, and that the
    /// ids written are the ones the gallery holds.</para>
    /// </remarks>
    [TestClass]
    public class PublishingSiteLiveTests : LiveTestBase
    {
        private const string AvailableWebTemplatesKey = "__WebTemplates";
        private const string InheritWebTemplatesKey = "__InheritWebTemplates";
        private const string AvailablePageLayoutsKey = "__PageLayouts";
        private const string DefaultPageLayoutKey = "__DefaultPageLayout";

        /// <summary>SharePoint Server Publishing - web scoped.</summary>
        private static readonly Guid WebPublishingFeature = new Guid("94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb");

        /// <summary>SharePoint Server Publishing Infrastructure - site collection scoped.</summary>
        private static readonly Guid SitePublishingFeature = new Guid("f6924d36-2fa8-4f0b-b16d-06b7250180fa");

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Handlers")]
        [TestCategory("Publishing")]
        [Timeout(30 * 60 * 1000)]
        public async Task Publishing_AppliesAndReExtractsOnASiteItCreates()
        {
            Uri siteUrl = null;

            try
            {
                using (PnPContext seed = await GetContextAsync().ConfigureAwait(false))
                {
                    string owner = (await seed.Web.GetCurrentUserAsync().ConfigureAwait(false)).LoginName;

                    // A fresh url every run: DeleteSiteCollectionAsync empties the recycle bin too,
                    // but a run that dies between create and delete would otherwise block the next.
                    siteUrl = new Uri($"https://{seed.Uri.DnsSafeHost}/sites/" +
                        $"pnpcoreprovisioningtestpub{Guid.NewGuid():N}");

                    Console.WriteLine($"Creating {siteUrl}");

                    using (PnPContext admin = await seed.GetSharePointAdmin()
                        .GetTenantAdminCenterContextAsync().ConfigureAwait(false))
                    {
                        // STS#0, not a modern template: the publishing infrastructure feature is a
                        // classic site collection feature and does not activate on a communication or
                        // group connected site.
                        var options = new ClassicSiteOptions(siteUrl, $"{TestPrefix}Publishing", "STS#0",
                            owner, Language.English, TimeZone.UTCPLUS0100_BRUSSELS_COPENHAGEN_MADRID_PARIS);

                        using (PnPContext created = await admin.GetSiteCollectionManager()
                            .CreateSiteCollectionAsync(options, new SiteCreationOptions
                            {
                                UsingApplicationPermissions = false,
                            }).ConfigureAwait(false))
                        {
                            Console.WriteLine($"Created {created.Uri}");
                        }

                        await AllowScriptingAsync(admin, siteUrl).ConfigureAwait(false);
                    }
                }

                await RunAgainstTheNewSiteAsync(siteUrl).ConfigureAwait(false);
            }
            finally
            {
                await DeleteSiteAsync(siteUrl).ConfigureAwait(false);
            }
        }

        private async Task RunAgainstTheNewSiteAsync(Uri siteUrl)
        {
            using (PnPContext seed = await GetContextAsync(1).ConfigureAwait(false))
            using (PnPContext context = await seed.CloneAsync(siteUrl).ConfigureAwait(false))
            {
                // ---------------------------------------------------------------------------------
                // 1. Web templates, with the handler activating publishing on the way in.
                // ---------------------------------------------------------------------------------
                Assert.IsFalse(await IsFeatureActiveAsync(context, WebPublishingFeature).ConfigureAwait(false),
                    "A brand new STS#0 site already had web publishing on - the activation path cannot be proven here.");

                var templatesTemplate = new ProvisioningTemplate
                {
                    Publishing = new PublishingModel
                    {
                        AutoCheckRequirements = AutoCheckRequirementsOptions.MakeCompliant,
                    },
                };

                templatesTemplate.Publishing.AvailableWebTemplates.Add(
                    new AvailableWebTemplate { LanguageCode = 1033, TemplateName = "STS#0" });
                templatesTemplate.Publishing.AvailableWebTemplates.Add(
                    new AvailableWebTemplate { LanguageCode = 1033, TemplateName = "BLANKINTERNET#0" });

                await ApplyAsync(context, templatesTemplate).ConfigureAwait(false);

                // MakeCompliant means the handler is responsible for turning publishing on, site
                // scope first. If it silently skipped instead, the property assertions below would
                // fail with no explanation - so this is checked on its own.
                //
                // Through a fresh context, because EnableAsync adds the feature to the collection it
                // was called on: asking the applying context would get that local addition back and
                // report success whether or not SharePoint did anything.
                using (PnPContext after = await seed.CloneAsync(siteUrl).ConfigureAwait(false))
                {
                    Assert.IsTrue(await IsSiteFeatureActiveAsync(after, SitePublishingFeature).ConfigureAwait(false),
                        "The site scoped publishing feature was not activated.");
                    Assert.IsTrue(await IsFeatureActiveAsync(after, WebPublishingFeature).ConfigureAwait(false),
                        "The web scoped publishing feature was not activated.");
                }

                Dictionary<string, string> written = await ReadPropertiesAsync(siteUrl).ConfigureAwait(false);

                Console.WriteLine($"{AvailableWebTemplatesKey} = {written[AvailableWebTemplatesKey]}");
                Console.WriteLine($"{InheritWebTemplatesKey} = {written[InheritWebTemplatesKey]}");

                XElement webTemplates = XElement.Parse(written[AvailableWebTemplatesKey]);
                Assert.AreEqual("webtemplates", webTemplates.Name.LocalName);

                CollectionAssert.AreEquivalent(
                    new[] { "STS#0", "BLANKINTERNET#0" },
                    webTemplates.Elements("lcid").Where(e => (string)e.Attribute("id") == "1033")
                        .Elements("webtemplate").Select(e => (string)e.Attribute("name")).ToList(),
                    "The web templates SharePoint stored are not the ones the template asked for.");

                // Without this the list above is ignored in favour of the parent's, so the site would
                // look provisioned and behave as though nothing was set.
                Assert.AreEqual("False", written[InheritWebTemplatesKey],
                    $"{InheritWebTemplatesKey} was not turned off, so the template list has no effect.");

                // ---------------------------------------------------------------------------------
                // 2. Page layouts, resolved through the master page gallery.
                // ---------------------------------------------------------------------------------
                // The publishing feature fills the gallery as part of activation, and activation can
                // outlive the call that requested it - so the gallery is waited for rather than
                // assumed. Bounded, with an explicit verdict either way.
                List<string> layoutNames = new List<string>();

                for (int attempt = 1; attempt <= 8; attempt++)
                {
                    using (PnPContext reader = await seed.CloneAsync(siteUrl).ConfigureAwait(false))
                    {
                        layoutNames = await ReadLayoutNamesAsync(reader).ConfigureAwait(false);
                    }

                    if (layoutNames.Count >= 2)
                    {
                        break;
                    }

                    await Task.Delay(TimeSpan.FromSeconds(15)).ConfigureAwait(false);
                }

                Console.WriteLine($"Gallery holds {layoutNames.Count} layouts, e.g. " +
                    string.Join(", ", layoutNames.Take(5)));

                if (layoutNames.Count < 2)
                {
                    Assert.Inconclusive("The master page gallery still held fewer than two page layouts " +
                        "two minutes after publishing was activated, so the layout path could not be " +
                        "exercised. The web templates half above did run.");
                }

                string first = layoutNames[0];
                string second = layoutNames[1];

                var layoutTemplate = new ProvisioningTemplate
                {
                    Publishing = new PublishingModel
                    {
                        AutoCheckRequirements = AutoCheckRequirementsOptions.MakeCompliant,
                    },
                };

                layoutTemplate.Publishing.PageLayouts.Add(new PageLayoutModel { Path = first });
                layoutTemplate.Publishing.PageLayouts.Add(new PageLayoutModel { Path = second, IsDefault = true });

                Console.WriteLine($"Asking for {layoutTemplate.Publishing.PageLayouts.Count} layout(s): " +
                    string.Join(", ", layoutTemplate.Publishing.PageLayouts.Select(l => $"{l.Path} (default={l.IsDefault})")));

                await ApplyAsync(context, layoutTemplate).ConfigureAwait(false);

                written = await ReadPropertiesAsync(siteUrl).ConfigureAwait(false);

                Console.WriteLine($"{AvailablePageLayoutsKey} = {written[AvailablePageLayoutsKey]}");
                Console.WriteLine($"{DefaultPageLayoutKey} = {written[DefaultPageLayoutKey]}");

                Assert.IsFalse(string.IsNullOrEmpty(written[AvailablePageLayoutsKey]),
                    $"'{AvailablePageLayoutsKey}' was not written. The apply reported no problem, so " +
                    "the handler either resolved no layouts or wrote them somewhere else.");

                XElement layouts = XElement.Parse(written[AvailablePageLayoutsKey]);
                Assert.AreEqual("pagelayouts", layouts.Name.LocalName);

                List<XElement> layoutElements = layouts.Elements("layout").ToList();
                Assert.AreEqual(2, layoutElements.Count, "Both layouts should have been resolved and written.");

                // The url is what a human reads, but the guid is what SharePoint matches on - a
                // handler that wrote a plausible url with the wrong id would produce a site whose
                // layout list is empty in the UI. So the ids are checked against the gallery.
                Dictionary<string, string> gallery = await ReadLayoutIdsAsync(context).ConfigureAwait(false);

                foreach (XElement layout in layoutElements)
                {
                    var url = (string)layout.Attribute("url");
                    var guid = (string)layout.Attribute("guid");

                    Assert.IsTrue(url.StartsWith("_catalogs/masterpage/", StringComparison.OrdinalIgnoreCase),
                        $"The layout url '{url}' is not site relative as the property bag expects.");

                    string name = url.Substring(url.LastIndexOf('/') + 1);

                    Assert.IsTrue(gallery.ContainsKey(name), $"'{name}' is not in the master page gallery.");
                    Assert.AreEqual(gallery[name], guid?.Trim('{', '}'),
                        $"The id written for '{name}' is not the one the gallery holds.");
                }

                Assert.AreEqual(second,
                    ((string)XElement.Parse(written[DefaultPageLayoutKey]).Attribute("url"))
                        .Split('/').Last(),
                    "The wrong layout was recorded as the default.");

                // ---------------------------------------------------------------------------------
                // 3. Read it all back through the handler.
                // ---------------------------------------------------------------------------------
                var configuration = new ExtractConfiguration();
                configuration.Handlers.Add(ConfigurationHandler.Publishing);

                // Fresh again: the applying context has the property bag materialised, so extracting
                // through it would read the engine's own in-memory copy and pass even if SharePoint
                // had stored nothing.
                ProvisioningTemplate extracted;

                using (PnPContext reader = await seed.CloneAsync(siteUrl).ConfigureAwait(false))
                {
                    extracted = await reader.GetProvisioningManager()
                        .GetTemplateAsync(configuration).ConfigureAwait(false);
                }

                Assert.IsNotNull(extracted.Publishing, "The extract produced no publishing element on a publishing site.");

                CollectionAssert.AreEquivalent(
                    new[] { "STS#0", "BLANKINTERNET#0" },
                    extracted.Publishing.AvailableWebTemplates.Select(t => t.TemplateName).ToList(),
                    "The extracted web templates do not match what was applied.");

                Assert.IsTrue(extracted.Publishing.AvailableWebTemplates.All(t => t.LanguageCode == 1033),
                    "The extracted web templates lost their language.");

                // Bare names, not gallery paths - that is what PnP Framework emits, and what the
                // template's Path means on the way back in.
                CollectionAssert.AreEquivalent(
                    new[] { first, second },
                    extracted.Publishing.PageLayouts.Select(l => l.Path).ToList(),
                    "The extracted page layout paths are not the bare names a template carries.");

                PageLayoutModel defaultLayout = extracted.Publishing.PageLayouts.SingleOrDefault(l => l.IsDefault);
                Assert.IsNotNull(defaultLayout, "No layout came back marked as the default.");
                Assert.AreEqual(second, defaultLayout.Path, "The wrong layout came back as the default.");
            }
        }

        #region Helpers

        /// <summary>
        /// Turns NoScript off on the site this test created.
        /// </summary>
        /// <remarks>
        /// <para><b>Every new site collection in SharePoint Online is NoScript</b>, classic ones
        /// included, and <c>ObjectPublishing</c> correctly declines to touch a NoScript site. So
        /// without this the test proves the guard - which is already proven elsewhere - and nothing
        /// about the write path. This is site setup, not the thing under test: the site is created
        /// and deleted by this method's caller, so nothing outside it is changed.</para>
        /// <para>The setting is stored tenant side and does not take effect immediately, hence the
        /// wait. It is a bounded wait with an explicit failure, not a blind sleep - if the site is
        /// still NoScript at the end, the test says so rather than going on to fail somewhere less
        /// obvious.</para>
        /// </remarks>
        private static async Task AllowScriptingAsync(PnPContext admin, Uri siteUrl)
        {
            ISiteCollectionProperties properties = await admin.GetSiteCollectionManager()
                .GetSiteCollectionPropertiesAsync(siteUrl).ConfigureAwait(false);

            properties.DenyAddAndCustomizePages = DenyAddAndCustomizePagesStatus.Disabled;
            await properties.UpdateAsync().ConfigureAwait(false);

            Console.WriteLine("Set DenyAddAndCustomizePages = Disabled, waiting for it to take effect.");

            for (int attempt = 1; attempt <= 12; attempt++)
            {
                using (PnPContext seed = await GetContextAsync(4).ConfigureAwait(false))
                using (PnPContext site = await seed.CloneAsync(siteUrl).ConfigureAwait(false))
                {
                    if (!await site.Web.IsNoScriptSiteAsync().ConfigureAwait(false))
                    {
                        Console.WriteLine($"Scripting allowed after {attempt} check(s).");
                        return;
                    }
                }

                await Task.Delay(TimeSpan.FromSeconds(15)).ConfigureAwait(false);
            }

            Assert.Inconclusive("The created site was still NoScript three minutes after scripting was " +
                "allowed on it, so the publishing write path could not be reached. This is a tenant " +
                "propagation delay rather than a defect in the handler - rerun the test.");
        }

        private static async Task ApplyAsync(PnPContext context, ProvisioningTemplate template)
        {
            var errors = new List<string>();

            await context.GetProvisioningManager().ApplyTemplateAsync(template, new ApplyConfiguration
            {
                MessagesDelegate = (message, type) =>
                {
                    Console.WriteLine($"[{type}] {message}");

                    if (type == ProvisioningMessageType.Warning || type == ProvisioningMessageType.Error)
                    {
                        errors.Add(message);
                    }
                },
            }).ConfigureAwait(false);

            // The handler reports and continues rather than throwing, so a silent skip would leave
            // the property assertions to fail with no clue why.
            Assert.AreEqual(0, errors.Count,
                $"The apply reported problems:{Environment.NewLine}{string.Join(Environment.NewLine, errors)}");
        }

        /// <summary>
        /// Reads the four property bag keys through a <b>fresh</b> context.
        /// </summary>
        /// <remarks>
        /// A fresh context on purpose: the applying context has the property bag materialised, so
        /// reading it back there would assert against the engine's own in-memory copy rather than
        /// against what SharePoint stored.
        /// </remarks>
        private static async Task<Dictionary<string, string>> ReadPropertiesAsync(Uri siteUrl)
        {
            using (PnPContext seed = await GetContextAsync(2).ConfigureAwait(false))
            using (PnPContext context = await seed.CloneAsync(siteUrl).ConfigureAwait(false))
            {
                await context.Web.LoadAsync(w => w.AllProperties).ConfigureAwait(false);

                var result = new Dictionary<string, string>();

                foreach (string key in new[]
                {
                    AvailableWebTemplatesKey, InheritWebTemplatesKey,
                    AvailablePageLayoutsKey, DefaultPageLayoutKey,
                })
                {
                    result[key] = context.Web.AllProperties.Values.TryGetValue(key, out object value)
                        ? value?.ToString()
                        : null;

                    Assert.IsFalse(string.IsNullOrEmpty(result[key]) && key == AvailableWebTemplatesKey,
                        $"'{key}' was not written at all.");
                }

                return result;
            }
        }

        /// <summary>
        /// The page layout file names in the master page gallery.
        /// </summary>
        /// <remarks>
        /// Read through the folder rather than through the list, which is how the handler does it.
        /// Reading it the same way would make this test agree with the handler by construction and
        /// prove nothing about whether the gallery really holds what was written.
        /// </remarks>
        private static async Task<List<string>> ReadLayoutNamesAsync(PnPContext context)
        {
            IFolder gallery = await GalleryFolderAsync(context).ConfigureAwait(false);

            return gallery.Files.AsRequested()
                .Select(f => f.Name)
                .Where(n => n.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase))
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static async Task<Dictionary<string, string>> ReadLayoutIdsAsync(PnPContext context)
        {
            IFolder gallery = await GalleryFolderAsync(context).ConfigureAwait(false);

            return gallery.Files.AsRequested()
                .GroupBy(f => f.Name, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First().UniqueId.ToString(), StringComparer.OrdinalIgnoreCase);
        }

        private static async Task<IFolder> GalleryFolderAsync(PnPContext context)
        {
            await context.Web.LoadAsync(w => w.ServerRelativeUrl).ConfigureAwait(false);

            return await context.Web.GetFolderByServerRelativeUrlAsync(
                $"{context.Web.ServerRelativeUrl.TrimEnd('/')}/_catalogs/masterpage",
                f => f.Files.QueryProperties(file => file.Name, file => file.UniqueId)).ConfigureAwait(false);
        }

        private static async Task<bool> IsFeatureActiveAsync(PnPContext context, Guid feature)
        {
            await context.Web.LoadAsync(w => w.Features.QueryProperties(f => f.DefinitionId)).ConfigureAwait(false);
            return context.Web.Features.AsRequested().Any(f => f.DefinitionId == feature);
        }

        private static async Task<bool> IsSiteFeatureActiveAsync(PnPContext context, Guid feature)
        {
            await context.Site.LoadAsync(s => s.Features.QueryProperties(f => f.DefinitionId)).ConfigureAwait(false);
            return context.Site.Features.AsRequested().Any(f => f.DefinitionId == feature);
        }

        /// <summary>
        /// Deletes the site, recycle bin included.
        /// </summary>
        /// <remarks>
        /// Never throws: a teardown failure would replace whatever the test actually found with a
        /// cleanup error. It is printed loudly instead, because a leaked site collection is real.
        /// </remarks>
        private static async Task DeleteSiteAsync(Uri siteUrl)
        {
            if (siteUrl == null)
            {
                return;
            }

            try
            {
                using (PnPContext context = await GetContextAsync(3).ConfigureAwait(false))
                {
                    await context.GetSiteCollectionManager()
                        .DeleteSiteCollectionAsync(siteUrl).ConfigureAwait(false);

                    Console.WriteLine($"Deleted {siteUrl}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"COULD NOT DELETE {siteUrl} - delete it by hand.{Environment.NewLine}{Describe(ex)}");
            }
        }

        #endregion
    }
}
