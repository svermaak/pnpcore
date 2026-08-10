using Microsoft.Extensions.Logging;
using PnP.Core.Provisioning.Model;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.ObjectHandlers.Utilities
{
    /// <summary>
    /// Decides whether a publishing-dependent handler may proceed, honouring the template's
    /// <c>AutoCheckRequirements</c>.
    /// </summary>
    /// <remarks>
    /// <para>Three handlers need this - <c>ObjectPublishing</c>, <c>ObjectImageRenditions</c> and
    /// <c>ObjectComposedLook</c> - and the semantics are subtle enough that three copies would
    /// drift. The template chooses:</para>
    /// <list type="table">
    /// <item><term>SkipIfNotCompliant</term><description>do nothing and say so</description></item>
    /// <item><term>MakeCompliant</term><description>activate the publishing features</description></item>
    /// <item><term>anything else</term><description>fail, because the template asserted a requirement that is not met</description></item>
    /// </list>
    /// <para><b>Order matters when activating.</b> The site-scoped feature has to come first: the
    /// web-scoped one depends on infrastructure the site feature installs, and activating them the
    /// other way round fails on a site that has neither.</para>
    /// <para><b>Activation outlives the request that asks for it.</b> On a site that has never had
    /// publishing, turning it on provisions galleries, layouts, master pages and a dozen lists, and
    /// takes longer than PnP Core's 100 second default HTTP timeout - so the call reliably fails
    /// while the work succeeds. Reporting that as "could not be activated" would make the very first
    /// apply to any non-publishing site skip its publishing settings, every time. See
    /// <c>ConfirmActivationAsync</c>.</para>
    /// </remarks>
    internal static class PublishingRequirements
    {
        /// <summary>SharePoint Server Publishing - web scoped.</summary>
        private static readonly Guid WebPublishingFeature = new Guid("94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb");

        /// <summary>SharePoint Server Publishing Infrastructure - site collection scoped.</summary>
        private static readonly Guid SitePublishingFeature = new Guid("f6924d36-2fa8-4f0b-b16d-06b7250180fa");

        /// <summary>
        /// Whether the web-scoped publishing feature is on.
        /// </summary>
        internal static async Task<bool> IsWebPublishingActiveAsync(PnPContext context)
        {
            try
            {
                await context.Web.LoadAsync(w => w.Features.QueryProperties(f => f.DefinitionId)).ConfigureAwait(false);
                return context.Web.Features.AsRequested().Any(f => f.DefinitionId == WebPublishingFeature);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static async Task<bool> IsSitePublishingActiveAsync(PnPContext context)
        {
            try
            {
                await context.Site.LoadAsync(s => s.Features.QueryProperties(f => f.DefinitionId)).ConfigureAwait(false);
                return context.Site.Features.AsRequested().Any(f => f.DefinitionId == SitePublishingFeature);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Ensures publishing is available, or reports why the caller should stop.
        /// </summary>
        /// <param name="context">The context</param>
        /// <param name="autoCheck">What the template asked for</param>
        /// <param name="reportWarning">Called when the caller should skip</param>
        /// <param name="what">What is being provisioned, for the message</param>
        /// <returns>Whether the caller may proceed</returns>
        /// <exception cref="InvalidOperationException">
        /// Publishing is off and the template neither asked to skip nor to make the site compliant.
        /// </exception>
        internal static async Task<bool> EnsureAsync(PnPContext context, AutoCheckRequirementsOptions autoCheck,
            Action<string> reportWarning, string what)
        {
            if (await IsWebPublishingActiveAsync(context).ConfigureAwait(false))
            {
                return true;
            }

            switch (autoCheck)
            {
                case AutoCheckRequirementsOptions.SkipIfNotCompliant:
                    string skipped = $"The publishing feature is not active on this web, so the {what} were skipped.";
                    context.Logger?.LogWarning("{Source}: {Message}", Constants.LOGGING_SOURCE, skipped);
                    reportWarning?.Invoke(skipped);
                    return false;

                case AutoCheckRequirementsOptions.MakeCompliant:
                    return await ActivateAsync(context, reportWarning, what).ConfigureAwait(false);

                default:
                    // The template said publishing is required and it is not there. Continuing would
                    // produce a site that silently lacks what the template promised.
                    throw new InvalidOperationException(
                        $"The publishing feature is not active on this web, so the {what} cannot be provisioned. " +
                        "Set AutoCheckRequirements to MakeCompliant or SkipIfNotCompliant to change this.");
            }
        }

        /// <summary>How long to keep waiting after an activation call gives up on the client side.</summary>
        /// <remarks>
        /// Turning on the publishing infrastructure provisions galleries, layouts, master pages and a
        /// dozen lists, and routinely takes longer than PnP Core's <b>100 second default HTTP
        /// timeout</b> (<c>PnPCore:HttpRequests:Timeout</c>). The request is abandoned by the client;
        /// SharePoint carries on regardless. Three minutes has been enough on a freshly created site.
        /// </remarks>
        private static readonly TimeSpan ActivationGracePeriod = TimeSpan.FromMinutes(3);

        private static readonly TimeSpan ActivationPollInterval = TimeSpan.FromSeconds(15);

        private static async Task<bool> ActivateAsync(PnPContext context, Action<string> reportWarning, string what)
        {
            try
            {
                // Site collection scope first - the web feature builds on it.
                if (!await IsSitePublishingActiveAsync(context).ConfigureAwait(false))
                {
                    await context.Site.Features.EnableAsync(SitePublishingFeature).ConfigureAwait(false);
                }

                await context.Web.Features.EnableAsync(WebPublishingFeature).ConfigureAwait(false);

                context.Logger?.LogInformation("{Source}: activated the publishing features for the {What}.",
                    Constants.LOGGING_SOURCE, what);

                return true;
            }
            catch (Exception ex)
            {
                // The call may have failed only on this side of the wire. Ask SharePoint what
                // actually happened before reporting a failure - see ConfirmActivationAsync.
                if (await ConfirmActivationAsync(context, ex).ConfigureAwait(false))
                {
                    return true;
                }

                string warning = $"The publishing features could not be activated, so the {what} were skipped: " +
                    ErrorText.Describe(ex);
                context.Logger?.LogWarning(ex, "{Source}: {Message}", Constants.LOGGING_SOURCE, warning);
                reportWarning?.Invoke(warning);
                return false;
            }
        }

        /// <summary>
        /// After a failed activation call, waits to see whether SharePoint completed it anyway.
        /// </summary>
        /// <remarks>
        /// <para><b>This is not a retry of a mutation.</b> It asks whether one that was already
        /// requested has landed. The distinction matters: activating the publishing infrastructure
        /// takes minutes on a cold site, so the client gives up long before the server does, and
        /// treating that as a failure means every first-time apply to a non-publishing site reports a
        /// warning and skips - while the site quietly becomes a publishing site a minute later.</para>
        /// <para><b>Every check runs on a cloned context.</b> <c>EnableAsync</c> adds the feature to
        /// the collection it was called on, so asking the original context would return that local
        /// addition and confirm an activation that never happened.</para>
        /// <para>The web-scoped feature is re-enabled if the site-scoped one arrived without it: that
        /// call is quick, and it is the ordering the site feature was blocking.</para>
        /// </remarks>
        private static async Task<bool> ConfirmActivationAsync(PnPContext context, Exception cause)
        {
            context.Logger?.LogInformation(
                "{Source}: the publishing feature activation call did not complete ({Reason}). " +
                "Waiting up to {Seconds}s to see whether SharePoint finished it anyway.",
                Constants.LOGGING_SOURCE, cause.Message, ActivationGracePeriod.TotalSeconds);

            DateTime deadline = DateTime.UtcNow.Add(ActivationGracePeriod);

            while (true)
            {
                try
                {
                    using (PnPContext fresh = await context.CloneAsync().ConfigureAwait(false))
                    {
                        if (await IsWebPublishingActiveAsync(fresh).ConfigureAwait(false))
                        {
                            context.Logger?.LogInformation(
                                "{Source}: the publishing features are active after all.", Constants.LOGGING_SOURCE);
                            return true;
                        }

                        if (await IsSitePublishingActiveAsync(fresh).ConfigureAwait(false))
                        {
                            // The slow half is done; the web feature is a quick call that the site
                            // feature was previously blocking.
                            await fresh.Web.Features.EnableAsync(WebPublishingFeature).ConfigureAwait(false);
                            return true;
                        }
                    }
                }
                catch (Exception)
                {
                    // Still settling. The deadline below is what ends this, not the first error.
                }

                if (DateTime.UtcNow >= deadline)
                {
                    return false;
                }

                await Task.Delay(ActivationPollInterval).ConfigureAwait(false);
            }
        }
    }
}
