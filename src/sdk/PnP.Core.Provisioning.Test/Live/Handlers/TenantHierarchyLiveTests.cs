using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Admin.Model.SharePoint;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.Provisioning.ObjectHandlers;
using PnP.Core.Provisioning.ObjectHandlers.Utilities;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using CommunicationSiteCollectionModel = PnP.Core.Provisioning.Model.CommunicationSiteCollection;
using TermGroupModel = PnP.Core.Provisioning.Model.TermGroup;
using ThemeModel = PnP.Core.Provisioning.Model.Theme;

namespace PnP.Core.Provisioning.Test.Live.Handlers
{
    /// <summary>
    /// Phase 8's exit gate: one tenant template that provisions tenant settings, a term group and a
    /// site collection, in one apply.
    /// </summary>
    /// <remarks>
    /// <para>Each hierarchy handler already has its own coverage. <b>What this adds is the ordering
    /// between them</b>, which is the part that cannot be tested one handler at a time:
    /// <c>ObjectHierarchyTenant</c> runs before <c>ObjectHierarchySequenceTermGroups</c>, which runs
    /// before <c>ObjectHierarchySequenceSites</c>, because a sequence's sites are routinely created
    /// from things the earlier handlers produce.</para>
    /// <para><b>Teams and Azure AD are deliberately left out.</b> A team leaves a Microsoft 365 group
    /// behind, and directory users consume real licence seats - neither is a throwaway artefact the
    /// way a site collection, a term group or a tenant theme is. They are covered by their own
    /// handlers' code paths and flagged in the backlog as needing an explicit decision.</para>
    /// </remarks>
    [TestClass]
    public class TenantHierarchyLiveTests : LiveTestBase
    {
        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Handlers")]
        [Timeout(30 * 60 * 1000)]
        public async Task Hierarchy_AppliesTenantSettingsThenTermGroupsThenSites()
        {
            string fixture = Guid.NewGuid().ToString("N").Substring(0, 12);

            string themeName = $"{TestPrefix}HierarchyTheme_{fixture}";
            string termGroupName = $"{TestPrefix}HierarchyGroup_{fixture}";
            var termGroupId = Guid.NewGuid();
            string listTitle = $"{TestPrefix}HierarchyList";

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
                    siteUrl = new Uri($"https://{context.Uri.DnsSafeHost}/sites/pnpcoreprovisioningtesthier{fixture}");

                    try
                    {
                        ProvisioningHierarchy hierarchy = BuildHierarchy(
                            siteUrl, themeName, termGroupName, termGroupId, listTitle);

                        var problems = new List<string>();

                        Console.WriteLine($"Applying hierarchy: theme '{themeName}', term group " +
                            $"'{termGroupName}', site {siteUrl}");

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
                            $"The hierarchy reported problems:{Environment.NewLine}{string.Join(Environment.NewLine, problems)}");

                        // 1. The tenant handler ran.
                        HashSet<string> themes = await TenantThemes.GetNamesAsync(admin).ConfigureAwait(false);
                        Assert.IsTrue(themes.Contains(themeName),
                            "ObjectHierarchyTenant did not apply the hierarchy's tenant settings.");

                        // 2. The term group handler ran - and kept the id the template gave it, which
                        // is the whole point of decision D7's CSOM-first taxonomy.
                        ITermGroup group = await admin.TermStore.Groups
                            .GetByIdAsync(termGroupId.ToString()).ConfigureAwait(false);

                        Assert.IsNotNull(group, "ObjectHierarchySequenceTermGroups did not create the term group.");
                        Assert.AreEqual(termGroupName, group.Name, "The term group was created with the wrong name.");

                        // 3. The sites handler ran, and its template was applied to the new site.
                        using (PnPContext seed = await GetContextAsync(1).ConfigureAwait(false))
                        using (PnPContext created = await seed.CloneAsync(siteUrl).ConfigureAwait(false))
                        {
                            await created.Web.LoadAsync(w => w.Lists.QueryProperties(l => l.Title)).ConfigureAwait(false);

                            Assert.IsTrue(created.Web.Lists.AsRequested().Any(l => l.Title == listTitle),
                                "ObjectHierarchySequenceSites created the site but its template was not applied.");
                        }

                        Console.WriteLine("All three hierarchy handlers ran, in order.");
                    }
                    finally
                    {
                        await CleanUpAsync(admin, siteUrl, themeName, termGroupId).ConfigureAwait(false);
                    }
                }
            }
        }

        private static ProvisioningHierarchy BuildHierarchy(Uri siteUrl, string themeName, string termGroupName,
            Guid termGroupId, string listTitle)
        {
            var hierarchy = new ProvisioningHierarchy
            {
                Tenant = new ProvisioningTenant(),
            };

            hierarchy.Tenant.Themes.Add(new ThemeModel
            {
                Name = themeName,
                IsInverted = false,
                Overwrite = true,
                Palette = "{\"themePrimary\":\"#0078d4\",\"themeDark\":\"#005a9e\",\"white\":\"#ffffff\"}",
            });

            var template = new ProvisioningTemplate { Id = "HIERARCHY-TEMPLATE" };

            template.Lists.Add(new ListInstance
            {
                Title = listTitle,
                Url = "Lists/PnPCoreProvisioningTestHierarchyList",
                TemplateType = (int)ListTemplateType.GenericList,
            });

            hierarchy.Templates.Add(template);

            // TermStore is null until something assigns one - in this port and in PnP Framework
            // alike, because "no term store element" and "an empty one" are different things in the
            // schema. The handler is null-safe; a caller building a hierarchy by hand has to say so.
            var sequence = new ProvisioningSequence
            {
                ID = "TENANTSEQUENCE",
                TermStore = new ProvisioningTermStore(),
            };

            sequence.TermStore.TermGroups.Add(new TermGroupModel
            {
                Id = termGroupId,
                Name = termGroupName,
                Description = "Created by the PnP Core provisioning tests",
            });

            var siteCollection = new CommunicationSiteCollectionModel
            {
                Url = siteUrl.ToString(),
                Title = $"{TestPrefix}Hierarchy",
                Description = "Created by the PnP Core provisioning tests",
                Language = 1033,
                ProvisioningId = "HIER-SITE",
            };

            siteCollection.Templates.Add("HIERARCHY-TEMPLATE");
            sequence.SiteCollections.Add(siteCollection);
            hierarchy.Sequences.Add(sequence);

            return hierarchy;
        }

        /// <summary>
        /// Removes everything the hierarchy created, whatever failed.
        /// </summary>
        /// <remarks>
        /// Three different lifetimes in one teardown - a site collection, a tenant-wide theme and a
        /// term group in the tenant term store - and none of them is swept by anything else.
        /// </remarks>
        private static async Task CleanUpAsync(PnPContext admin, Uri siteUrl, string themeName, Guid termGroupId)
        {
            try
            {
                await TenantThemes.DeleteAsync(admin, themeName).ConfigureAwait(false);
                Console.WriteLine($"Deleted theme {themeName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"COULD NOT DELETE THEME {themeName}: {Describe(ex)}");
            }

            try
            {
                await DeleteTermGroupDeepAsync(admin, termGroupId.ToString()).ConfigureAwait(false);
                Console.WriteLine($"Deleted term group {termGroupId}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"COULD NOT DELETE TERM GROUP {termGroupId}: {Describe(ex)}");
            }

            if (siteUrl == null)
            {
                return;
            }

            try
            {
                using (PnPContext context = await GetContextAsync(2).ConfigureAwait(false))
                {
                    ISiteCollectionManager manager = context.GetSiteCollectionManager();

                    if (await manager.SiteExistsAsync(siteUrl).ConfigureAwait(false))
                    {
                        await manager.DeleteSiteCollectionAsync(siteUrl).ConfigureAwait(false);
                        Console.WriteLine($"Deleted {siteUrl}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"COULD NOT DELETE {siteUrl} - delete it by hand.{Environment.NewLine}{Describe(ex)}");
            }
        }
    }
}
