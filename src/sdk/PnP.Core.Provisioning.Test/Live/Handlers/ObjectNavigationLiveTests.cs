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
using NavigationModel = PnP.Core.Provisioning.Model.Navigation;
using NavigationNodeModel = PnP.Core.Provisioning.Model.NavigationNode;

namespace PnP.Core.Provisioning.Test.Live.Handlers
{
    /// <summary>
    /// Live coverage for <c>ObjectNavigation</c>, including backlog <b>T9</b>.
    /// </summary>
    /// <remarks>
    /// <para>The quick launch is <b>restored</b>, not swept: it is site configuration that always
    /// exists, so a test that replaces it has to put the original nodes back. Every test here
    /// records the current nodes first and rebuilds them in a <c>finally</c>.</para>
    /// <para>The nesting test earns its place. PnP Core's <c>INavigationNode</c> has no children
    /// collection, so an extract that went through the model alone would flatten every submenu and
    /// still look like it worked - the assertion is on the child, one level down.</para>
    /// </remarks>
    [TestClass]
    public class ObjectNavigationLiveTests : LiveTestBase
    {
        private static string ParentTitle => $"{TestPrefix}Parent";

        private static string ChildTitle => $"{TestPrefix}Child";

        #region Apply

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Handlers")]
        public async Task Navigation_CreatesQuickLaunchNodesIncludingAChild()
        {
            using (PnPContext context = await GetNoGroupContextAsync().ConfigureAwait(false))
            {
                List<NavigationNodeModel> original = await ReadQuickLaunchAsync(context).ConfigureAwait(false);
                Console.WriteLine($"Quick launch had {original.Count} node(s)");

                try
                {
                    var parent = new NavigationNodeModel
                    {
                        Title = ParentTitle,
                        Url = "https://example.com/parent",
                        IsExternal = true,
                    };

                    parent.NavigationNodes.Add(new NavigationNodeModel
                    {
                        Title = ChildTitle,
                        Url = "https://example.com/child",
                        IsExternal = true,
                    });

                    var structural = new StructuralNavigation { RemoveExistingNodes = true };
                    structural.NavigationNodes.Add(parent);

                    var template = new ProvisioningTemplate
                    {
                        Navigation = new NavigationModel(
                            null,
                            new CurrentNavigation(CurrentNavigationType.Structural, structural, null),
                            null),
                    };

                    await context.GetProvisioningManager().ApplyTemplateAsync(template, Reporting()).ConfigureAwait(false);

                    using (PnPContext fresh = await GetNoGroupContextAsync(1).ConfigureAwait(false))
                    {
                        var configuration = new ExtractConfiguration();
                        configuration.Handlers.Add(ConfigurationHandler.Navigation);

                        ProvisioningTemplate extracted = await fresh.GetProvisioningManager()
                            .GetTemplateAsync(configuration).ConfigureAwait(false);

                        StructuralNavigation current = extracted.Navigation?.CurrentNavigation?.StructuralNavigation;
                        Assert.IsNotNull(current, "The current navigation was not extracted.");

                        Console.WriteLine($"Extracted: {string.Join(", ", current.NavigationNodes.Select(n => n.Title))}");

                        NavigationNodeModel extractedParent = current.NavigationNodes
                            .FirstOrDefault(n => n.Title == ParentTitle);

                        Assert.IsNotNull(extractedParent, "The parent node was not created.");

                        // RemoveExistingNodes was set, so the suite's node should be the only one.
                        Assert.AreEqual(1, current.NavigationNodes.Count,
                            "RemoveExistingNodes did not clear the quick launch before adding.");

                        // The child is the assertion that matters: PnP Core's node model has no
                        // children, so an extract that used it alone would report a flat menu here
                        // and pass every other check in this test.
                        Assert.AreEqual(1, extractedParent.NavigationNodes.Count,
                            "The child node was not created, or nesting was lost on extract.");

                        Assert.AreEqual(ChildTitle, extractedParent.NavigationNodes[0].Title);
                    }
                }
                finally
                {
                    await RestoreQuickLaunchAsync(original).ConfigureAwait(false);
                }
            }
        }

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Handlers")]
        public async Task Navigation_AppliedTwiceDoesNotAccumulateNodes()
        {
            using (PnPContext context = await GetNoGroupContextAsync().ConfigureAwait(false))
            {
                List<NavigationNodeModel> original = await ReadQuickLaunchAsync(context).ConfigureAwait(false);

                try
                {
                    IProvisioningManager manager = context.GetProvisioningManager();

                    await manager.ApplyTemplateAsync(BuildTemplate(), Reporting()).ConfigureAwait(false);
                    await manager.ApplyTemplateAsync(BuildTemplate(), Reporting()).ConfigureAwait(false);

                    using (PnPContext fresh = await GetNoGroupContextAsync(1).ConfigureAwait(false))
                    {
                        List<NavigationNodeModel> nodes = await ReadQuickLaunchAsync(fresh).ConfigureAwait(false);
                        Console.WriteLine($"After two applies: {nodes.Count} node(s)");

                        // The count is the assertion. This is the failure mode the site footer had
                        // in phase 5 - each apply appending to the last, unnoticed because "the node
                        // is present" stayed true the whole time.
                        Assert.AreEqual(1, nodes.Count, "The second apply appended instead of replacing.");
                    }
                }
                finally
                {
                    await RestoreQuickLaunchAsync(original).ConfigureAwait(false);
                }
            }
        }

        #endregion

        #region T9 - managed metadata navigation

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Handlers")]
        public async Task Navigation_ManagedMetadataIsRefusedByNameRatherThanSilentlyIgnored()
        {
            using (PnPContext context = await GetNoGroupContextAsync().ConfigureAwait(false))
            {
                List<NavigationNodeModel> original = await ReadQuickLaunchAsync(context).ConfigureAwait(false);
                string warning = null;

                try
                {
                    var template = new ProvisioningTemplate
                    {
                        Navigation = new NavigationModel(
                            null,
                            new CurrentNavigation(CurrentNavigationType.Managed, null,
                                new ManagedNavigation
                                {
                                    TermStoreId = "11111111-1111-1111-1111-111111111111",
                                    TermSetId = "22222222-2222-2222-2222-222222222222",
                                }),
                            null),
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

                    Assert.IsNotNull(warning, "Managed metadata navigation was skipped without saying so.");

                    // T9 is a real gap. The warning has to name the term set, so a template author
                    // can tell which part of their template did nothing - a generic "not supported"
                    // leaves them guessing.
                    StringAssert.Contains(warning, "22222222-2222-2222-2222-222222222222",
                        "The warning does not name the term set that was skipped.");

                    using (PnPContext fresh = await GetNoGroupContextAsync(1).ConfigureAwait(false))
                    {
                        List<NavigationNodeModel> nodes = await ReadQuickLaunchAsync(fresh).ConfigureAwait(false);

                        // Refusing must leave the site alone. Clearing the nodes and then declining
                        // would be worse than doing nothing.
                        Assert.AreEqual(original.Count, nodes.Count,
                            "Refusing managed navigation still changed the site's quick launch.");
                    }
                }
                finally
                {
                    await RestoreQuickLaunchAsync(original).ConfigureAwait(false);
                }
            }
        }

        #endregion

        #region Helpers

        private static ProvisioningTemplate BuildTemplate()
        {
            var structural = new StructuralNavigation { RemoveExistingNodes = true };

            structural.NavigationNodes.Add(new NavigationNodeModel
            {
                Title = ParentTitle,
                Url = "https://example.com/parent",
                IsExternal = true,
            });

            return new ProvisioningTemplate
            {
                Navigation = new NavigationModel(
                    null,
                    new CurrentNavigation(CurrentNavigationType.Structural, structural, null),
                    null),
            };
        }

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

        private static async Task<List<NavigationNodeModel>> ReadQuickLaunchAsync(PnPContext context)
        {
            var nodes = new List<NavigationNodeModel>();

            await context.Web.LoadAsync(w => w.Navigation.QueryProperties(n => n.QuickLaunch)).ConfigureAwait(false);

            foreach (INavigationNode node in context.Web.Navigation.QuickLaunch.AsRequested())
            {
                nodes.Add(new NavigationNodeModel
                {
                    Title = node.Title,
                    Url = node.Url,
                    IsExternal = node.IsExternal,
                });
            }

            return nodes;
        }

        /// <summary>
        /// Rebuilds the quick launch as the test found it.
        /// </summary>
        /// <remarks>
        /// Only the top level is restored - the read above does not walk children, and a partial
        /// restore that invented a hierarchy would be worse than a flat one. The suite's own nodes
        /// are what matter for isolation, and those are removed by the rebuild.
        /// </remarks>
        private static async Task RestoreQuickLaunchAsync(List<NavigationNodeModel> original)
        {
            try
            {
                using (PnPContext context = await GetNoGroupContextAsync(2).ConfigureAwait(false))
                {
                    await context.Web.LoadAsync(w => w.Navigation.QueryProperties(n => n.QuickLaunch))
                        .ConfigureAwait(false);

                    await context.Web.Navigation.QuickLaunch.DeleteAllNodesAsync().ConfigureAwait(false);

                    foreach (NavigationNodeModel node in original)
                    {
                        await context.Web.Navigation.QuickLaunch.AddAsync(new NavigationNodeOptions
                        {
                            Title = node.Title,
                            Url = node.Url,
                        }).ConfigureAwait(false);
                    }

                    Console.WriteLine($"Restored {original.Count} quick launch node(s).");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"COULD NOT RESTORE the quick launch: {Describe(ex)}");
            }
        }

        #endregion
    }
}
