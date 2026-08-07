using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.Provisioning.ObjectHandlers;
using PnP.Core.Services;
using System;
using System.Linq;
using System.Threading.Tasks;
using PropertyBagEntryModel = PnP.Core.Provisioning.Model.PropertyBagEntry;

namespace PnP.Core.Provisioning.Test.Live.Handlers
{
    /// <summary>
    /// Live coverage for <c>ObjectPropertyBagEntry</c>.
    /// </summary>
    /// <remarks>
    /// Property bag writes are refused outright on a NoScript site, which the configured test site
    /// may well be. Where that is so the tests report it rather than failing - and the first test
    /// says which case applies, so a skip is never ambiguous.
    /// </remarks>
    [TestClass]
    public class PropertyBagLiveTests : LiveTestBase
    {
        private static string TestKey => $"{TestPrefix}Key";

        private static async Task<bool> SkipIfNoScriptAsync(PnPContext context)
        {
            if (await context.Web.IsNoScriptSiteAsync().ConfigureAwait(false))
            {
                Assert.Inconclusive(
                    "This is a NoScript site, so property bag writes are refused by SharePoint and the handler " +
                    "correctly skips them. Run against a non-NoScript site to verify the write path.");
                return true;
            }

            return false;
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Handlers")]
        public async Task PropertyBag_ExtractReadsTheWebPropertyBag()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                Console.WriteLine($"NoScript site: {await context.Web.IsNoScriptSiteAsync().ConfigureAwait(false)}");

                ProvisioningTemplate template = await context.GetProvisioningManager()
                    .GetTemplateAsync().ConfigureAwait(false);

                Assert.IsTrue(template.PropertyBagEntries.Count > 0,
                    "Every web has property bag entries - an empty result means extraction did not read them.");

                Console.WriteLine($"Property bag entries extracted: {template.PropertyBagEntries.Count}");
                Console.WriteLine($"Marked as indexed              : {template.PropertyBagEntries.Count(e => e.Indexed)}");
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Handlers")]
        public async Task PropertyBag_RoundTripsAnEntry()
        {
            using (PnPContext context = await GetClassicContextAsync().ConfigureAwait(false))
            {
                if (await SkipIfNoScriptAsync(context).ConfigureAwait(false)) return;

                string value = $"value-{DateTime.UtcNow:HHmmss}";

                try
                {
                    var template = new ProvisioningTemplate();
                    template.PropertyBagEntries.Add(new PropertyBagEntryModel
                    {
                        Key = TestKey,
                        Value = value,
                        Overwrite = true,
                    });

                    await context.GetProvisioningManager().ApplyTemplateAsync(template).ConfigureAwait(false);

                    using (PnPContext fresh = await GetClassicContextAsync(1).ConfigureAwait(false))
                    {
                        IWeb web = await fresh.Web.GetAsync(w => w.AllProperties).ConfigureAwait(false);
                        Assert.AreEqual(value, web.AllProperties.GetString(TestKey, null),
                            "The property bag entry did not round-trip.");
                    }
                }
                finally
                {
                    await RemoveTestKeyAsync(context).ConfigureAwait(false);
                }
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Handlers")]
        public async Task PropertyBag_IndexedEntryIsMarkedIndexedAndExtractsThatWay()
        {
            using (PnPContext context = await GetClassicContextAsync().ConfigureAwait(false))
            {
                if (await SkipIfNoScriptAsync(context).ConfigureAwait(false)) return;

                try
                {
                    var template = new ProvisioningTemplate();
                    template.PropertyBagEntries.Add(new PropertyBagEntryModel
                    {
                        Key = TestKey,
                        Value = "indexed-value",
                        Overwrite = true,
                        Indexed = true,
                    });

                    await context.GetProvisioningManager().ApplyTemplateAsync(template).ConfigureAwait(false);

                    // Re-extract through the engine: this exercises the base64/Unicode decoding of
                    // vti_indexedpropertykeys, which has no getter on IWeb and so is the part most
                    // likely to be subtly wrong.
                    using (PnPContext fresh = await GetClassicContextAsync(1).ConfigureAwait(false))
                    {
                        ProvisioningTemplate extracted = await fresh.GetProvisioningManager()
                            .GetTemplateAsync().ConfigureAwait(false);

                        PropertyBagEntryModel entry = extracted.PropertyBagEntries
                            .FirstOrDefault(e => e.Key == TestKey);

                        Assert.IsNotNull(entry, "The entry was not extracted.");
                        Assert.IsTrue(entry.Indexed,
                            "The entry was written as indexed but did not extract that way - the " +
                            "vti_indexedpropertykeys decoding is wrong.");
                    }
                }
                finally
                {
                    await RemoveTestKeyAsync(context).ConfigureAwait(false);
                }
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Handlers")]
        public async Task PropertyBag_OverwriteFalseDoesNotReplaceAnExistingValue()
        {
            using (PnPContext context = await GetClassicContextAsync().ConfigureAwait(false))
            {
                if (await SkipIfNoScriptAsync(context).ConfigureAwait(false)) return;

                try
                {
                    // Seed a value...
                    var seed = new ProvisioningTemplate();
                    seed.PropertyBagEntries.Add(new PropertyBagEntryModel { Key = TestKey, Value = "original", Overwrite = true });
                    await context.GetProvisioningManager().ApplyTemplateAsync(seed).ConfigureAwait(false);

                    // ...then apply a template that must NOT replace it. Overwrite=false means
                    // "create if absent", and getting that backwards would have a re-applied
                    // template quietly reset values a site owner had changed.
                    var second = new ProvisioningTemplate();
                    second.PropertyBagEntries.Add(new PropertyBagEntryModel { Key = TestKey, Value = "replacement", Overwrite = false });
                    await context.GetProvisioningManager().ApplyTemplateAsync(second).ConfigureAwait(false);

                    using (PnPContext fresh = await GetClassicContextAsync(1).ConfigureAwait(false))
                    {
                        IWeb web = await fresh.Web.GetAsync(w => w.AllProperties).ConfigureAwait(false);
                        Assert.AreEqual("original", web.AllProperties.GetString(TestKey, null),
                            "Overwrite=false replaced an existing value - it must only create absent ones.");
                    }
                }
                finally
                {
                    await RemoveTestKeyAsync(context).ConfigureAwait(false);
                }
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Handlers")]
        public async Task PropertyBag_SystemPropertiesAreSkippedUnlessExplicitlyAllowed()
        {
            using (PnPContext context = await GetClassicContextAsync().ConfigureAwait(false))
            {
                if (await SkipIfNoScriptAsync(context).ConfigureAwait(false)) return;

                // A key with a system prefix must be skipped by default. Writing to vti_* can
                // destabilise a site, so the default has to be "refuse".
                var template = new ProvisioningTemplate();
                template.PropertyBagEntries.Add(new PropertyBagEntryModel
                {
                    Key = "vti_" + TestPrefix + "ShouldNotBeWritten",
                    Value = "nope",
                    Overwrite = true,
                });

                await context.GetProvisioningManager().ApplyTemplateAsync(template).ConfigureAwait(false);

                using (PnPContext fresh = await GetClassicContextAsync(1).ConfigureAwait(false))
                {
                    IWeb web = await fresh.Web.GetAsync(w => w.AllProperties).ConfigureAwait(false);
                    Assert.IsNull(web.AllProperties.GetString("vti_" + TestPrefix + "ShouldNotBeWritten", null),
                        "A system-prefixed property was written without OverwriteSystemPropertyBagValues.");
                }
            }
        }

        private static async Task RemoveTestKeyAsync(PnPContext context)
        {
            try
            {
                IWeb web = await context.Web.GetAsync(w => w.AllProperties).ConfigureAwait(false);

                if (web.AllProperties.Values.ContainsKey(TestKey))
                {
                    await web.RemoveIndexedPropertyAsync(TestKey).ConfigureAwait(false);
                    web.AllProperties[TestKey] = string.Empty;
                    await web.AllProperties.UpdateAsync().ConfigureAwait(false);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Could not clean up property bag key {TestKey}: {Describe(ex)}");
            }
        }
    }
}
