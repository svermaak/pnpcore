using Microsoft.VisualStudio.TestTools.UnitTesting;
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
using ComposedLookModel = PnP.Core.Provisioning.Model.ComposedLook;
using PublishingModel = PnP.Core.Provisioning.Model.Publishing;

namespace PnP.Core.Provisioning.Test.Live.Handlers
{
    /// <summary>
    /// Live coverage for <c>ObjectPublishing</c> (<b>T17</b>) and <c>ObjectComposedLook</c>
    /// (<b>T19</b>).
    /// </summary>
    /// <remarks>
    /// <para>Both are classic-only, and both are stored as <b>XML in the web's property bag</b>
    /// rather than through an API. That makes the property bag the thing worth asserting on: the
    /// SharePoint UI reads those exact keys, so a handler that writes almost-right XML produces a
    /// site that looks provisioned and behaves as though nothing happened.</para>
    /// <para>The web-templates test is the one that can run anywhere - it needs publishing to be
    /// active to reach the handler at all, so on a site without it the test asserts the refusal
    /// instead. Which of the two it does is decided from the site, not assumed.</para>
    /// </remarks>
    [TestClass]
    public class PublishingAndComposedLookLiveTests : LiveTestBase
    {
        private const string AvailableWebTemplatesKey = "__WebTemplates";
        private const string ComposedLookInfoKey = "_PnP_ProvisioningTemplateComposedLookInfo";

        #region Publishing

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Handlers")]
        public async Task Publishing_WritesTheAvailableWebTemplatesAsTheXmlSharePointReads()
        {
            using (PnPContext context = await GetClassicContextAsync().ConfigureAwait(false))
            {
                bool publishingActive = await IsWebPublishingActiveAsync(context).ConfigureAwait(false);
                Console.WriteLine($"Web publishing active: {publishingActive}");

                string original = await ReadPropertyAsync(context, AvailableWebTemplatesKey).ConfigureAwait(false);

                try
                {
                    var template = new ProvisioningTemplate
                    {
                        Publishing = new PublishingModel
                        {
                            AutoCheckRequirements = AutoCheckRequirementsOptions.SkipIfNotCompliant,
                        },
                    };

                    template.Publishing.AvailableWebTemplates.Add(new AvailableWebTemplate
                    {
                        LanguageCode = 1033,
                        TemplateName = "STS#0",
                    });
                    template.Publishing.AvailableWebTemplates.Add(new AvailableWebTemplate
                    {
                        LanguageCode = 1033,
                        TemplateName = "BLANKINTERNET#0",
                    });

                    string warning = null;

                    await context.GetProvisioningManager().ApplyTemplateAsync(template, new ApplyConfiguration
                    {
                        MessagesDelegate = (message, type) =>
                        {
                            Console.WriteLine($"[{type}] {message}");
                            if (type == ProvisioningMessageType.Warning)
                            {
                                warning = message;
                            }
                        },
                    }).ConfigureAwait(false);

                    if (!publishingActive)
                    {
                        // The handler must decline and say so rather than writing settings the site
                        // cannot honour. That is the assertion on this site.
                        Assert.IsNotNull(warning, "Publishing is off and the handler neither applied nor reported.");
                        return;
                    }

                    using (PnPContext fresh = await GetClassicContextAsync(1).ConfigureAwait(false))
                    {
                        string xml = await ReadPropertyAsync(fresh, AvailableWebTemplatesKey).ConfigureAwait(false);
                        Console.WriteLine($"{AvailableWebTemplatesKey} = {xml}");

                        Assert.IsFalse(string.IsNullOrEmpty(xml), "No web templates XML was written.");

                        // Asserting on the parsed shape, not on a string match: SharePoint reads
                        // this, so what matters is that the elements and attributes are the ones it
                        // expects, grouped by LCID.
                        XElement root = XElement.Parse(xml);
                        Assert.AreEqual("webtemplates", root.Name.LocalName);

                        XElement lcid = root.Elements("lcid").FirstOrDefault(e => (string)e.Attribute("id") == "1033");
                        Assert.IsNotNull(lcid, "The templates were not grouped under an <lcid id=\"1033\"> element.");

                        List<string> names = lcid.Elements("webtemplate")
                            .Select(e => (string)e.Attribute("name")).ToList();

                        CollectionAssert.AreEquivalent(new[] { "STS#0", "BLANKINTERNET#0" }, names,
                            "The web template names were not written as SharePoint expects them.");
                    }
                }
                finally
                {
                    await RestorePropertyAsync(AvailableWebTemplatesKey, original).ConfigureAwait(false);
                }
            }
        }

        #endregion

        #region Composed look

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Handlers")]
        public async Task ComposedLook_IsSkippedOnANoScriptSite()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                if (!await IsNoScriptAsync(context).ConfigureAwait(false))
                {
                    Assert.Inconclusive("This site is not NoScript, so the refusal path cannot be exercised on it.");
                }

                string warning = null;

                var template = new ProvisioningTemplate
                {
                    ComposedLook = new ComposedLookModel
                    {
                        Name = $"{TestPrefix}Look",
                        ColorFile = "{themecatalog}/15/Palette001.spcolor",
                    },
                };

                await context.GetProvisioningManager().ApplyTemplateAsync(template, new ApplyConfiguration
                {
                    MessagesDelegate = (message, type) =>
                    {
                        Console.WriteLine($"[{type}] {message}");
                        if (type == ProvisioningMessageType.Warning)
                        {
                            warning = message;
                        }
                    },
                }).ConfigureAwait(false);

                Assert.IsNotNull(warning, "A composed look on a NoScript site was skipped without saying so.");
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Handlers")]
        public async Task ComposedLook_ExtractReportsCurrentWhenNothingWasRecorded()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                string recorded = await ReadPropertyAsync(context, ComposedLookInfoKey).ConfigureAwait(false);

                if (!string.IsNullOrEmpty(recorded))
                {
                    Assert.Inconclusive("This site already has a recorded composed look, so the fallback cannot be exercised.");
                }

                var configuration = new ExtractConfiguration();
                configuration.Handlers.Add(ConfigurationHandler.ComposedLook);

                ProvisioningTemplate extracted = await context.GetProvisioningManager()
                    .GetTemplateAsync(configuration).ConfigureAwait(false);

                Assert.IsNotNull(extracted.ComposedLook, "No composed look element was produced at all.");
                Console.WriteLine($"Composed look: {extracted.ComposedLook.Name}");

                // SharePoint keeps the resolved theme file urls but not the name of the look that
                // produced them, so "Current" is the only honest answer when nothing was recorded.
                // Inventing a name would make a round trip apply something that was never chosen.
                Assert.AreEqual("Current", extracted.ComposedLook.Name,
                    "With nothing recorded, the extract should report the site's current look rather than a name.");
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Handlers")]
        public async Task ComposedLook_AppliesAndRoundTripsItsNameThroughThePropertyBag()
        {
            using (PnPContext context = await GetClassicContextAsync().ConfigureAwait(false))
            {
                if (await IsNoScriptAsync(context).ConfigureAwait(false))
                {
                    Assert.Inconclusive("The classic test site is NoScript, so composed looks cannot be applied to it.");
                }

                string original = await ReadPropertyAsync(context, ComposedLookInfoKey).ConfigureAwait(false);

                try
                {
                    var template = new ProvisioningTemplate
                    {
                        ComposedLook = new ComposedLookModel
                        {
                            // A look that exists in every classic site's design catalog, so the
                            // handler resolves its files from there rather than needing uploads.
                            Name = "Office",
                        },
                    };

                    await context.GetProvisioningManager().ApplyTemplateAsync(template, Reporting()).ConfigureAwait(false);

                    using (PnPContext fresh = await GetClassicContextAsync(1).ConfigureAwait(false))
                    {
                        var configuration = new ExtractConfiguration();
                        configuration.Handlers.Add(ConfigurationHandler.ComposedLook);

                        ProvisioningTemplate extracted = await fresh.GetProvisioningManager()
                            .GetTemplateAsync(configuration).ConfigureAwait(false);

                        Console.WriteLine($"Extracted composed look: {extracted.ComposedLook?.Name}");

                        // The name is the whole point: SharePoint does not store it, so recovering
                        // it proves the property bag round trip works. Without that an
                        // extract-then-apply cycle silently loses which look was chosen.
                        Assert.AreEqual("Office", extracted.ComposedLook?.Name,
                            "The applied look's name was not recovered from the property bag.");
                    }
                }
                finally
                {
                    await RestorePropertyAsync(ComposedLookInfoKey, original, classic: true).ConfigureAwait(false);
                }
            }
        }

        #endregion

        #region Helpers

        private static ApplyConfiguration Reporting()
        {
            return new ApplyConfiguration
            {
                MessagesDelegate = (message, type) =>
                {
                    if (type == ProvisioningMessageType.Warning || type == ProvisioningMessageType.Error)
                    {
                        Console.WriteLine($"[{type}] {message}");
                    }
                },
            };
        }

        private static async Task<bool> IsWebPublishingActiveAsync(PnPContext context)
        {
            try
            {
                await context.Web.LoadAsync(w => w.Features.QueryProperties(f => f.DefinitionId)).ConfigureAwait(false);

                return context.Web.Features.AsRequested()
                    .Any(f => f.DefinitionId == new Guid("94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb"));
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static async Task<string> ReadPropertyAsync(PnPContext context, string key)
        {
            try
            {
                await context.Web.LoadAsync(w => w.AllProperties).ConfigureAwait(false);

                return context.Web.AllProperties.Values.TryGetValue(key, out object value)
                    ? value?.ToString()
                    : null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Could not read '{key}': {Describe(ex)}");
                return null;
            }
        }

        /// <summary>
        /// Puts a property bag entry back the way the test found it.
        /// </summary>
        /// <remarks>
        /// These entries are site configuration, not artefacts a sweep can delete. An empty string
        /// is the closest thing to "absent" the property bag offers - removing a key is not
        /// something the API exposes.
        /// </remarks>
        private static async Task RestorePropertyAsync(string key, string original, bool classic = true)
        {
            try
            {
                using (PnPContext context = classic
                    ? await GetClassicContextAsync(2).ConfigureAwait(false)
                    : await GetContextAsync(2).ConfigureAwait(false))
                {
                    await context.Web.LoadAsync(w => w.AllProperties).ConfigureAwait(false);

                    context.Web.AllProperties[key] = original ?? string.Empty;
                    await context.Web.AllProperties.UpdateAsync().ConfigureAwait(false);

                    Console.WriteLine($"Restored '{key}'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"COULD NOT RESTORE '{key}': {Describe(ex)}");
            }
        }

        #endregion
    }
}
