using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Services.Core.CSOM;
using PnP.Core.Provisioning.Services.Core.CSOM.Requests.UserResources;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.Test.Live.Csom
{
    /// <summary>
    /// Live round trips for the two user resource CSOM requests (<b>T6</b>).
    /// </summary>
    /// <remarks>
    /// <para><b>The phase 6 blocker.</b> PnP Core has <c>IUserResource</c> support commented out
    /// across <c>IContentType</c>, <c>IField</c>, <c>IUserCustomAction</c> and <c>IWeb</c> - a known
    /// deferral. A dozen handlers carry localizable <c>Title</c> and <c>Description</c> values, so
    /// decision <b>D6</b> put T6 at the front of phase 6 and made these a required phase 4
    /// deliverable.</para>
    /// <para>Each test writes a localized value, reads it back, and restores what was there before.
    /// Localization writes are additive - setting a value for one culture does not disturb another -
    /// so the blast radius is one property in one language.</para>
    /// </remarks>
    [TestClass]
    public class UserResourceRequestLiveTests : LiveTestBase
    {
        /// <summary>
        /// A culture the test site is very unlikely to be running in, so writing to it cannot
        /// change what anyone actually sees.
        /// </summary>
        private const string ProbeCulture = "fi-FI";

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("UserResources")]
        public async Task SetAndGetValueForUICulture_RoundTripsOnTheWebTitle()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                (Guid siteId, Guid webId) = await CsomRequestSender.GetSiteAndWebIdAsync(context).ConfigureAwait(false);

                var resource = UserResourcePath.ForWeb(siteId, webId, ResourceProperty.Title);
                string original = null;

                try
                {
                    original = await CsomRequestSender.SendAsync(context,
                        new GetValueForUICultureRequest(resource, ProbeCulture)).ConfigureAwait(false);

                    string localized = $"{TestPrefix}Otsikko_{DateTime.UtcNow:HHmmss}";

                    await CsomRequestSender.SendAsync(context,
                        new SetValueForUICultureRequest(resource, ProbeCulture, localized)).ConfigureAwait(false);

                    string readBack = await CsomRequestSender.SendAsync(context,
                        new GetValueForUICultureRequest(resource, ProbeCulture)).ConfigureAwait(false);

                    Assert.AreEqual(localized, readBack,
                        "The localized web title did not round-trip. T6 is what a dozen phase 6 handlers depend on.");
                }
                finally
                {
                    // Put back whatever was there - typically the site's default title.
                    if (original != null)
                    {
                        try
                        {
                            await CsomRequestSender.SendAsync(context,
                                new SetValueForUICultureRequest(resource, ProbeCulture, original)).ConfigureAwait(false);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Could not restore the original {ProbeCulture} web title: {ex.Message}");
                        }
                    }
                }
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("UserResources")]
        public async Task SetAndGetValueForUICulture_RoundTripsOnAListTitle()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                (Guid siteId, Guid webId) = await CsomRequestSender.GetSiteAndWebIdAsync(context).ConfigureAwait(false);

                string listTitle = $"{TestPrefix}Loc_{DateTime.UtcNow:HHmmssfff}";
                IList list = null;

                try
                {
                    list = await context.Web.Lists.AddAsync(listTitle, ListTemplateType.GenericList).ConfigureAwait(false);

                    // A list identity, not a web one - this is the path ObjectListInstance will use.
                    var resource = UserResourcePath.ForList(siteId, webId, list.Id, ResourceProperty.Title);

                    string localized = $"{TestPrefix}Lista_{DateTime.UtcNow:HHmmss}";

                    await CsomRequestSender.SendAsync(context,
                        new SetValueForUICultureRequest(resource, ProbeCulture, localized)).ConfigureAwait(false);

                    string readBack = await CsomRequestSender.SendAsync(context,
                        new GetValueForUICultureRequest(resource, ProbeCulture)).ConfigureAwait(false);

                    Assert.AreEqual(localized, readBack, "The localized list title did not round-trip.");
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
        [TestCategory("UserResources")]
        public async Task SetValueForUICulture_DoesNotDisturbOtherCultures()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                (Guid siteId, Guid webId) = await CsomRequestSender.GetSiteAndWebIdAsync(context).ConfigureAwait(false);

                string listTitle = $"{TestPrefix}MultiLoc_{DateTime.UtcNow:HHmmssfff}";
                IList list = null;

                try
                {
                    list = await context.Web.Lists.AddAsync(listTitle, ListTemplateType.GenericList).ConfigureAwait(false);
                    var resource = UserResourcePath.ForList(siteId, webId, list.Id, ResourceProperty.Title);

                    // A template localizing into several languages sends one request per culture.
                    // They must not overwrite each other - which is the whole point of the API.
                    await CsomRequestSender.SendAsync(context,
                        new SetValueForUICultureRequest(resource, "fi-FI", "Suomeksi")).ConfigureAwait(false);
                    await CsomRequestSender.SendAsync(context,
                        new SetValueForUICultureRequest(resource, "de-DE", "Auf Deutsch")).ConfigureAwait(false);

                    Assert.AreEqual("Suomeksi", await CsomRequestSender.SendAsync(context,
                        new GetValueForUICultureRequest(resource, "fi-FI")).ConfigureAwait(false));

                    Assert.AreEqual("Auf Deutsch", await CsomRequestSender.SendAsync(context,
                        new GetValueForUICultureRequest(resource, "de-DE")).ConfigureAwait(false));
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
        [TestCategory("UserResources")]
        public async Task SetValueForUICulture_BatchesSeveralCulturesInOneRoundTrip()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                (Guid siteId, Guid webId) = await CsomRequestSender.GetSiteAndWebIdAsync(context).ConfigureAwait(false);

                string listTitle = $"{TestPrefix}BatchLoc_{DateTime.UtcNow:HHmmssfff}";
                IList list = null;

                try
                {
                    list = await context.Web.Lists.AddAsync(listTitle, ListTemplateType.GenericList).ConfigureAwait(false);
                    var resource = UserResourcePath.ForList(siteId, webId, list.Id, ResourceProperty.Title);

                    // A template localizing into eight languages should cost one round trip, not
                    // eight. This is what CsomRequestSender.SendManyAsync exists for.
                    var batch = new System.Collections.Generic.List<PnP.Core.Services.Core.CSOM.Requests.IRequest<object>>
                    {
                        new SetValueForUICultureRequest(resource, "fi-FI", "Yksi"),
                        new SetValueForUICultureRequest(resource, "de-DE", "Zwei"),
                        new SetValueForUICultureRequest(resource, "fr-FR", "Trois"),
                    };

                    await CsomRequestSender.SendManyAsync(context, batch).ConfigureAwait(false);

                    Assert.AreEqual("Yksi", await CsomRequestSender.SendAsync(context,
                        new GetValueForUICultureRequest(resource, "fi-FI")).ConfigureAwait(false));
                    Assert.AreEqual("Zwei", await CsomRequestSender.SendAsync(context,
                        new GetValueForUICultureRequest(resource, "de-DE")).ConfigureAwait(false));
                    Assert.AreEqual("Trois", await CsomRequestSender.SendAsync(context,
                        new GetValueForUICultureRequest(resource, "fr-FR")).ConfigureAwait(false));
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
    }
}
