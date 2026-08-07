using Microsoft.Extensions.Logging;
using PnP.Core.Provisioning.Extensibility;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.Provisioning.ObjectHandlers.TokenDefinitions;
using PnP.Core.Services;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Calls out to the custom handlers a template - or the caller - supplies.
    /// </summary>
    /// <remarks>
    /// <para>The engine's escape hatch: anything the schema cannot express is done by an assembly
    /// the template names. Handlers come from three places and are unioned, exactly as PnP Framework
    /// did it - the template's <c>ExtensibilityHandlers</c>, its deprecated <c>Providers</c>, and
    /// whatever the caller passed in the configuration.</para>
    /// <para><b>Failures are not swallowed.</b> A handler that throws aborts the run. That is
    /// deliberate: an extensibility handler is doing work no other handler can, so continuing past
    /// it would produce a site that looks provisioned and is not.</para>
    /// </remarks>
    internal class ObjectExtensibilityHandlers : ObjectHandlerBase
    {
        private readonly ExtensibilityManager extensibilityManager = new ExtensibilityManager();

        public override string Name => "Extensibility Providers";

        public override string InternalName => "ExtensibilityProviders";

        public override bool WillProvision(PnPContext context, ProvisioningTemplate template, ApplyConfiguration configuration)
        {
            _willProvision ??= GetHandlersToApply(template, configuration).Any();
            return _willProvision.Value;
        }

        public override bool WillExtract(PnPContext context, ProvisioningTemplate template, ExtractConfiguration configuration)
        {
            _willExtract ??= configuration?.ToCreationInformation()?.ExtensibilityHandlers?.Any() ?? false;
            return _willExtract.Value;
        }

        /// <summary>
        /// Runs the custom token providers and folds the tokens they return into the parser.
        /// </summary>
        /// <remarks>
        /// Called by <see cref="ProvisioningManager"/> immediately after the parser is built and
        /// <b>before</b> the first handler runs - a custom token has to exist before anything can
        /// reference it. It is a separate entry point rather than part of
        /// <see cref="ProvisionObjectsAsync"/> for exactly that reason.
        /// </remarks>
        internal async Task<TokenParser> AddExtendedTokensAsync(PnPContext context, ProvisioningTemplate template, TokenParser parser, ApplyConfiguration configuration)
        {
            using (context.Logger?.BeginScope(Name))
            {
                IEnumerable<ExtensibilityHandler> handlers = configuration != null
                    ? template.ExtensibilityHandlers.Union(configuration.ToApplyingInformation().ExtensibilityHandlers)
                    : template.ExtensibilityHandlers;

                foreach (ExtensibilityHandler handler in handlers.Where(h => h.Enabled))
                {
                    if (!string.IsNullOrEmpty(handler.Configuration))
                    {
                        handler.Configuration = parser.ParseString(handler.Configuration);
                    }

                    context.Logger?.LogDebug("{Source}: calling token provider extensibility callout {Assembly}",
                        Constants.LOGGING_SOURCE, handler.Assembly);

                    IEnumerable<TokenDefinition> providedTokens = await extensibilityManager
                        .ExecuteTokenProviderCallOutAsync(context, handler, template).ConfigureAwait(false);

                    if (providedTokens == null)
                    {
                        continue;
                    }

                    foreach (TokenDefinition token in providedTokens)
                    {
                        parser.AddToken(token);
                    }
                }

                return parser;
            }
        }

        public override async Task<TokenParser> ProvisionObjectsAsync(PnPContext context, ProvisioningTemplate template, TokenParser parser, ApplyConfiguration configuration)
        {
            using (context.Logger?.BeginScope(Name))
            {
                List<ExtensibilityHandler> handlersToProcess = GetHandlersToApply(template, configuration).ToList();

                if (handlersToProcess.Count == 0)
                {
                    return parser;
                }

                int currentCount = 0;
                foreach (ExtensibilityHandler handler in handlersToProcess.Where(h => h.Enabled))
                {
                    currentCount++;

                    if (!string.IsNullOrEmpty(handler.Configuration))
                    {
                        handler.Configuration = parser.ParseString(handler.Configuration);
                    }

                    context.Logger?.LogDebug("{Source}: calling extensibility callout {Assembly}",
                        Constants.LOGGING_SOURCE, handler.Assembly);

                    WriteSubProgress("Extensibility handler", handler.Type, currentCount, handlersToProcess.Count);

                    await extensibilityManager.ExecuteExtensibilityProvisionCallOutAsync(
                        context, handler, template, configuration, parser, context.Logger).ConfigureAwait(false);
                }

                WriteMessage("Done processing extensibility handlers", ProvisioningMessageType.Completed);

                return parser;
            }
        }

        public override async Task<ProvisioningTemplate> ExtractObjectsAsync(PnPContext context, ProvisioningTemplate template, ExtractConfiguration configuration)
        {
            using (context.Logger?.BeginScope(Name))
            {
                ProvisioningTemplateCreationInformation creationInformation = configuration?.ToCreationInformation();
                if (creationInformation?.ExtensibilityHandlers == null)
                {
                    return template;
                }

                foreach (ExtensibilityHandler handler in creationInformation.ExtensibilityHandlers.Where(h => h.Enabled))
                {
                    context.Logger?.LogDebug("{Source}: calling extensibility callout {Assembly}",
                        Constants.LOGGING_SOURCE, handler.Assembly);

                    template = await extensibilityManager.ExecuteExtensibilityExtractionCallOutAsync(
                        context, handler, template, configuration, context.Logger).ConfigureAwait(false);
                }

                return template;
            }
        }

        /// <summary>
        /// The handlers to run for an apply: the template's own, its deprecated
        /// <c>Providers</c> list, and any the caller supplied.
        /// </summary>
        private static IEnumerable<ExtensibilityHandler> GetHandlersToApply(ProvisioningTemplate template, ApplyConfiguration configuration)
        {
#pragma warning disable CS0618 // Providers is obsolete but templates written against older schemas still use it
            IEnumerable<ExtensibilityHandler> handlers = template.ExtensibilityHandlers.Union(template.Providers);
#pragma warning restore CS0618

            IEnumerable<ExtensibilityHandler> fromConfiguration = configuration?.ToApplyingInformation()?.ExtensibilityHandlers;

            return fromConfiguration != null ? handlers.Union(fromConfiguration) : handlers;
        }
    }
}
