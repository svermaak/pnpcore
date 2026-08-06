using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Test.Utilities;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.Test.Live
{
    /// <summary>
    /// Shared helpers for the live tests.
    /// </summary>
    /// <remarks>
    /// <para>These need a real tenant, configured through <c>env.txt</c> plus
    /// <c>appsettings.&lt;env&gt;.json</c>. They are all <c>[TestCategory("Live")]</c> so the
    /// offline suite - and CI, which has no tenant - never picks them up.</para>
    /// <para><b>Everything created here is named with a common prefix and torn down.</b> The
    /// taxonomy tests in particular write to the <em>tenant-wide</em> term store, which is not
    /// scoped to a test site and which nothing else cleans up, so teardown runs in a
    /// <c>finally</c> rather than at the end of the happy path.</para>
    /// </remarks>
    public abstract class LiveTestBase
    {
        /// <summary>
        /// Prefix for every artefact these tests create, so a leaked one is obvious and can be
        /// swept up by hand.
        /// </summary>
        protected const string TestPrefix = "PnPCoreProvisioningTest_";

        /// <summary>
        /// Skips the test with a clear reason rather than failing it, for the capabilities a
        /// tenant may legitimately not have - SP2013 workflows, publishing, site policies.
        /// </summary>
        /// <remarks>
        /// <b>Always reports the underlying error.</b> "Capability unavailable" is a conclusion, and
        /// stating it without the evidence hides real defects behind a green-looking skip - which
        /// happened once already: <c>AddWebPart</c> was reported as "classic web parts unavailable"
        /// while <c>GetWebPartDefinitions</c> was succeeding through the very same web part manager.
        /// Print the detail so the reader can judge whether the skip is honest.
        /// </remarks>
        protected static void SkipIfUnavailable(string capability, Exception ex)
        {
            Assert.Inconclusive(
                $"{capability} appears to be unavailable on this tenant, so this request could not be verified." +
                $"{Environment.NewLine}Judge for yourself from the error below - if a *related* request succeeded, " +
                $"this is a defect rather than a missing capability.{Environment.NewLine}{Environment.NewLine}" +
                $"{Describe(ex)}");
        }

        /// <summary>
        /// Renders an exception with the detail SharePoint actually returned.
        /// </summary>
        /// <remarks>
        /// PnP Core's <c>ServiceException.Message</c> is a generic banner ("CSOM service exception").
        /// The useful text - what SharePoint said and why - is on the <c>ServiceError</c> and in the
        /// inner exception chain.
        /// </remarks>
        protected static string Describe(Exception ex)
        {
            var lines = new List<string>();

            if (ex is ServiceException serviceException && serviceException.Error is ServiceError error)
            {
                lines.Add($"HTTP {error.HttpResponseCode}, code '{error.Code}'");
                if (!string.IsNullOrEmpty(error.Message))
                {
                    lines.Add(error.Message);
                }
            }

            for (Exception current = ex; current != null; current = current.InnerException)
            {
                lines.Add($"{current.GetType().Name}: {current.Message}");
            }

            return string.Join(Environment.NewLine, lines);
        }

        /// <summary>
        /// Gets a context against the configured test site.
        /// </summary>
        protected static async Task<PnPContext> GetContextAsync(int id = 0)
        {
            return await TestCommon.Instance.GetContextAsync(TestCommon.TestSite, id).ConfigureAwait(false);
        }

        /// <summary>
        /// Gets a context against the configured classic (STS#0) test site.
        /// </summary>
        /// <remarks>
        /// <para>Needed by anything that <b>writes</b> classic artefacts. A modern group-connected
        /// site is <b>NoScript</b> by default, which removes the "Add and Customize Pages"
        /// permission - so <c>ImportWebPart</c> and <c>AddWebPart</c> fail there with a permissions
        /// error while <c>GetWebPartDefinitions</c> succeeds through the very same web part
        /// manager. That asymmetry is easy to misread as "the request is broken".</para>
        /// <para>It is also exactly why the engine has <c>NoScriptTemplateCleaner</c>: applying a
        /// template full of classic artefacts to a modern site cannot work, and the engine strips
        /// them rather than failing.</para>
        /// </remarks>
        protected static async Task<PnPContext> GetClassicContextAsync(int id = 0)
        {
            return await TestCommon.Instance.GetContextAsync(TestCommon.ClassicSTS0TestSite, id).ConfigureAwait(false);
        }

        /// <summary>
        /// Reports whether the context's web is a NoScript site.
        /// </summary>
        /// <remarks>
        /// Used to tell "this request is broken" apart from "this site does not permit it", which
        /// are otherwise indistinguishable from the error alone.
        /// </remarks>
        protected static async Task<bool> IsNoScriptAsync(PnPContext context)
        {
            try
            {
                return await context.Web.IsNoScriptSiteAsync().ConfigureAwait(false);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Deletes a term group, emptying it first.
        /// </summary>
        /// <remarks>
        /// <b>A group containing term sets cannot be deleted.</b> Graph refuses with a generic
        /// service exception, and a teardown that only tries the group leaves it behind - which is
        /// exactly what an earlier version of this helper did, leaking seven groups into the
        /// tenant-wide term store before the hygiene check caught it. The sets have to go first.
        /// </remarks>
        protected static async Task DeleteTermGroupDeepAsync(ITermGroup group)
        {
            if (group == null)
            {
                return;
            }

            // Empty the group before removing it.
            try
            {
                await group.LoadAsync(g => g.Sets.QueryProperties(s => s.Id)).ConfigureAwait(false);

                foreach (ITermSet set in group.Sets.AsRequested().ToList())
                {
                    try
                    {
                        await set.DeleteAsync().ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"  could not delete term set {set.Id}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  could not enumerate sets in group {group.Id}: {ex.Message}");
            }

            await group.DeleteAsync().ConfigureAwait(false);
        }

        /// <summary>
        /// Deletes a term group by id, emptying it first, tolerating its absence.
        /// </summary>
        protected static async Task DeleteTermGroupDeepAsync(PnPContext context, string groupId)
        {
            if (string.IsNullOrEmpty(groupId))
            {
                return;
            }

            try
            {
                ITermGroup group = await context.TermStore.Groups.GetByIdAsync(groupId).ConfigureAwait(false);
                await DeleteTermGroupDeepAsync(group).ConfigureAwait(false);
            }
            catch (Exception)
            {
                // Already gone, or not visible - the prefix sweep reports anything real.
            }
        }

        /// <summary>
        /// Removes every term group this test suite has left behind.
        /// </summary>
        /// <remarks>
        /// Belt and braces: each test cleans up after itself, but a crashed run can leave a group
        /// in the tenant term store. Calling this at the start of a run sweeps those up.
        /// </remarks>
        protected static async Task CleanUpLeakedTermGroupsAsync(PnPContext context)
        {
            try
            {
                await context.TermStore.LoadAsync(t => t.Groups.QueryProperties(g => g.Id, g => g.Name)).ConfigureAwait(false);

                foreach (ITermGroup group in context.TermStore.Groups.AsRequested()
                    .Where(g => g.Name != null && g.Name.StartsWith(TestPrefix, StringComparison.Ordinal))
                    .ToList())
                {
                    try
                    {
                        await DeleteTermGroupDeepAsync(group).ConfigureAwait(false);
                    }
                    catch (Exception)
                    {
                        // Leave it for the hygiene test to report by name.
                    }
                }
            }
            catch (Exception)
            {
                // No term store access - the taxonomy tests will report that themselves.
            }
        }
    }
}
