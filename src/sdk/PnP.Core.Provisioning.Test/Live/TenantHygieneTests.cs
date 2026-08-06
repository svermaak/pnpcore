using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.Test.Live
{
    /// <summary>
    /// Checks the live suite has not left anything behind in the tenant.
    /// </summary>
    /// <remarks>
    /// <para>The taxonomy tests write to the <b>tenant-wide</b> term store, not to a disposable
    /// test site. Nothing else cleans that up, and a crashed run can leave a term group behind. So
    /// this is worth having as an explicit, runnable check rather than an assumption.</para>
    /// <para>Run it after a live session. It reports what it finds and removes it.</para>
    /// </remarks>
    [TestClass]
    public class TenantHygieneTests : LiveTestBase
    {
        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Hygiene")]
        public async Task NoTestTermGroupsAreLeftBehind()
        {
            using (PnPContext context = await GetContextAsync().ConfigureAwait(false))
            {
                await context.TermStore.LoadAsync(t => t.Groups.QueryProperties(g => g.Id, g => g.Name)).ConfigureAwait(false);

                List<ITermGroup> leaked = context.TermStore.Groups.AsRequested()
                    .Where(g => g.Name != null && g.Name.StartsWith(TestPrefix, StringComparison.Ordinal))
                    .ToList();

                if (leaked.Count == 0)
                {
                    Console.WriteLine("Term store is clean - no test groups left behind.");
                    return;
                }

                Console.WriteLine($"Found {leaked.Count} leaked test term group(s); removing:");

                var failures = new List<string>();
                foreach (ITermGroup group in leaked)
                {
                    Console.WriteLine($"  {group.Name} ({group.Id})");
                    try
                    {
                        // Deep delete: a group with term sets in it refuses to go.
                        await DeleteTermGroupDeepAsync(group).ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        // A group with term sets still in it can refuse deletion. Report it by name
                        // so it can be removed by hand rather than left silently.
                        failures.Add($"{group.Name} ({group.Id}): {ex.Message}");
                    }
                }

                Assert.AreEqual(0, failures.Count,
                    "These test term groups could not be removed automatically and need deleting by hand " +
                    $"in the SharePoint admin centre:{Environment.NewLine}{string.Join(Environment.NewLine, failures)}");

                Console.WriteLine($"Removed {leaked.Count} leaked test term group(s).");
            }
        }
    }
}
