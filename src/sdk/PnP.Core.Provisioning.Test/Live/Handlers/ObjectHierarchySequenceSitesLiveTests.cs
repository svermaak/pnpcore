using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Admin.Model.SharePoint;
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
using CommunicationSiteCollectionModel = PnP.Core.Provisioning.Model.CommunicationSiteCollection;

namespace PnP.Core.Provisioning.Test.Live.Handlers
{
    /// <summary>
    /// Live coverage for <c>ObjectHierarchySequenceSites</c> - backlog <b>T14</b>.
    /// </summary>
    /// <remarks>
    /// <para>A sequence is the one part of the engine that <b>creates site collections</b>, so there
    /// is no way to exercise it without making one. The site is named with the suite prefix, created
    /// by this test, and deleted in a <c>finally</c> - recycle bin included - in the same shape as
    /// the publishing test.</para>
    /// <para><b>One communication site, not a team site.</b> A group connected site would leave a
    /// Microsoft 365 group behind whose deletion is a separate, slower operation, and it would prove
    /// nothing extra about the handler: the branch that differs is which options object is built, and
    /// that is decided by a switch this test does not need to walk twice.</para>
    /// <para>The assertion that matters is not "a site exists" - it is that the template named by the
    /// sequence was applied <em>to the new site</em>, and that the site's tokens resolved. A handler
    /// that created the site and silently skipped the template would pass a weaker test.</para>
    /// </remarks>
    [TestClass]
    public class ObjectHierarchySequenceSitesLiveTests : LiveTestBase
    {
        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Handlers")]
        [Timeout(30 * 60 * 1000)]
        public async Task Sequence_CreatesTheSiteAndAppliesTheTemplateItNames()
        {
            Uri siteUrl = null;

            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                PnPContext admin;

                try
                {
                    admin = await context.GetSharePointAdmin().GetTenantAdminCenterContextAsync().ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    SkipIfUnavailable("The tenant admin site", ex);
                    return;
                }

                using (admin)
                {
                    string fixture = Guid.NewGuid().ToString("N").Substring(0, 12);
                    siteUrl = new Uri($"https://{context.Uri.DnsSafeHost}/sites/pnpcoreprovisioningtestseq{fixture}");

                    string listTitle = $"{TestPrefix}SequenceList";

                    try
                    {
                        var hierarchy = new ProvisioningHierarchy();

                        // A template with one unmistakable artefact. The list is what proves the
                        // template reached the new site rather than the site merely existing.
                        var template = new ProvisioningTemplate { Id = "SEQUENCE-TEMPLATE" };

                        template.Lists.Add(new ListInstance
                        {
                            Title = listTitle,
                            Url = "Lists/PnPCoreProvisioningTestSequenceList",
                            TemplateType = (int)ListTemplateType.GenericList,

                            // Through a token that only resolves if the sequence published it, so a
                            // template applied against the wrong site or with a stale parser is
                            // visible in the artefact rather than silently fine.
                            Description = "Created by sequence {sequencesiteurl:SEQ-SITE}",
                        });

                        hierarchy.Templates.Add(template);

                        var sequence = new ProvisioningSequence { ID = "TENANTSEQUENCE" };

                        var siteCollection = new CommunicationSiteCollectionModel
                        {
                            Url = siteUrl.ToString(),
                            Title = $"{TestPrefix}Sequence",
                            Description = "Created by the PnP Core provisioning tests",
                            Language = 1033,
                            ProvisioningId = "SEQ-SITE",
                        };

                        siteCollection.Templates.Add("SEQUENCE-TEMPLATE");
                        sequence.SiteCollections.Add(siteCollection);
                        hierarchy.Sequences.Add(sequence);

                        Console.WriteLine($"Applying sequence, target {siteUrl}");

                        var problems = new List<string>();

                        await admin.GetProvisioningManager().ApplyTenantTemplateAsync(
                            hierarchy, "TENANTSEQUENCE", new ApplyConfiguration
                            {
                                MessagesDelegate = (message, type) =>
                                {
                                    Console.WriteLine($"[{type}] {message}");

                                    if (type == ProvisioningMessageType.Warning || type == ProvisioningMessageType.Error)
                                    {
                                        problems.Add(message);
                                    }
                                },
                            }).ConfigureAwait(false);

                        Assert.AreEqual(0, problems.Count,
                            $"The sequence reported problems:{Environment.NewLine}{string.Join(Environment.NewLine, problems)}");

                        // Everything below reads through a fresh context bound to the new site: the
                        // handler's own context has the site materialised, so asserting through it
                        // would check the engine's in-memory copy.
                        using (PnPContext seed = await GetContextAsync(1).ConfigureAwait(false))
                        using (PnPContext created = await seed.CloneAsync(siteUrl).ConfigureAwait(false))
                        {
                            IWeb web = await created.Web.GetAsync(w => w.Title, w => w.Url).ConfigureAwait(false);

                            Console.WriteLine($"Created {web.Url}, title '{web.Title}'");
                            Assert.AreEqual($"{TestPrefix}Sequence", web.Title, "The site was created with the wrong title.");

                            await created.Web.LoadAsync(w => w.Lists.QueryProperties(l => l.Title, l => l.Description))
                                .ConfigureAwait(false);

                            IList list = created.Web.Lists.AsRequested()
                                .FirstOrDefault(l => l.Title == listTitle);

                            Assert.IsNotNull(list,
                                "The sequence created the site but its template was not applied to it - " +
                                $"no list called '{listTitle}' exists.");

                            Console.WriteLine($"List description: {list.Description}");

                            // The token had to resolve to this very site's url, which only the
                            // sequence could have supplied.
                            StringAssert.Contains(list.Description, siteUrl.ToString().TrimEnd('/'),
                                "The {sequencesiteurl:SEQ-SITE} token did not resolve to the created site, " +
                                "so the template was applied with a parser that did not know about it.");
                        }
                    }
                    finally
                    {
                        await DeleteSiteAsync(siteUrl).ConfigureAwait(false);
                    }
                }
            }
        }

        /// <summary>
        /// Deletes the site, recycle bin included, without ever throwing.
        /// </summary>
        private static async Task DeleteSiteAsync(Uri siteUrl)
        {
            if (siteUrl == null)
            {
                return;
            }

            try
            {
                using (PnPContext context = await GetContextAsync(2).ConfigureAwait(false))
                {
                    ISiteCollectionManager manager = context.GetSiteCollectionManager();

                    if (!await manager.SiteExistsAsync(siteUrl).ConfigureAwait(false))
                    {
                        Console.WriteLine($"{siteUrl} was never created, nothing to delete.");
                        return;
                    }

                    await manager.DeleteSiteCollectionAsync(siteUrl).ConfigureAwait(false);
                    Console.WriteLine($"Deleted {siteUrl}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"COULD NOT DELETE {siteUrl} - delete it by hand.{Environment.NewLine}{Describe(ex)}");
            }
        }
    }
}
