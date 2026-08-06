using Microsoft.Extensions.Logging;
using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.Provisioning.ObjectHandlers.Utilities;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Runs the provisioning engine: builds the handler list, orders it, threads the token parser
    /// through it, and fires the webhooks a template asks for.
    /// </summary>
    /// <remarks>
    /// The PnP Core equivalent of PnP Framework's <c>SiteToTemplateConversion</c>. Two behaviours
    /// here are load-bearing and were preserved exactly - see <see cref="BuildApplyHandlers"/> for
    /// the three-pass ordering, and the handler loop in <c>ApplyTemplateAsync</c> for the parser
    /// chain.
    /// </remarks>
    internal sealed class ProvisioningManager : IProvisioningManager
    {
        private readonly PnPContext context;

        internal ProvisioningManager(PnPContext pnpContext)
        {
            context = pnpContext ?? throw new ArgumentNullException(nameof(pnpContext));
        }

        #region Apply a template

        /// <inheritdoc/>
        public async Task ApplyTemplateAsync(ProvisioningTemplate template, ApplyConfiguration configuration = null)
        {
            await ApplyTemplateAsync(template, configuration, false, null).ConfigureAwait(false);
        }

        /// <inheritdoc/>
        public void ApplyTemplate(ProvisioningTemplate template, ApplyConfiguration configuration = null)
        {
            ApplyTemplateAsync(template, configuration).GetAwaiter().GetResult();
        }

        internal async Task<TokenParser> ApplyTemplateAsync(ProvisioningTemplate template, ApplyConfiguration configuration,
            bool calledFromHierarchy, TokenParser tokenParser)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            ILogger logger = context.Logger;
            using (logger?.BeginScope(PnPCoreProvisioningResources.Provisioning_ObjectHandlers_Provisioning))
            {
                ProvisioningProgressDelegate progressDelegate = null;
                ProvisioningMessagesDelegate messagesDelegate = null;
                ProvisioningSiteProvisionedDelegate siteProvisionedDelegate = null;

                if (configuration != null)
                {
                    // Apply the configuration parameters to the template
                    ApplyConfigurationParameters(configuration.Parameters, template.Parameters);

                    progressDelegate = configuration.ProgressDelegate;
                    messagesDelegate = configuration.MessagesDelegate;
                    siteProvisionedDelegate = configuration.SiteProvisionedDelegate;
                }
                else
                {
                    // When no configuration was passed we want to execute all handlers
                    configuration = new ApplyConfiguration();
                }

                ProvisioningTemplateApplyingInformation applyingInformation = configuration.ToApplyingInformation();

                IWeb web = await context.Web.GetAsync(w => w.Url, w => w.Title, w => w.ServerRelativeUrl).ConfigureAwait(false);
                await context.Site.LoadAsync(s => s.ServerRelativeUrl).ConfigureAwait(false);

                // Check the template's scope against the target, when one was declared
                if (template.Scope == ProvisioningTemplateScope.RootSite && ObjectHandlerBase.IsSubSite(web))
                {
                    logger?.LogError("{Source}: {Message}", Constants.LOGGING_SOURCE,
                        PnPCoreProvisioningResources.SiteToTemplateConversion_ScopeOfTemplateDoesNotMatchTarget);
                    throw new InvalidOperationException(PnPCoreProvisioningResources.SiteToTemplateConversion_ScopeOfTemplateDoesNotMatchTarget);
                }

                CultureInfo currentCultureInfoValue = System.Threading.Thread.CurrentThread.CurrentCulture;
                if (!string.IsNullOrEmpty(template.TemplateCultureInfo))
                {
                    System.Threading.Thread.CurrentThread.CurrentCulture = int.TryParse(template.TemplateCultureInfo, out int cultureInfoValue)
                        ? new CultureInfo(cultureInfoValue)
                        : new CultureInfo(template.TemplateCultureInfo);
                }

                try
                {
                    // Warn when the target site does not share a base template with the template's source
                    await WarnOnAsymmetricBaseTemplatesAsync(web, template, messagesDelegate).ConfigureAwait(false);

                    List<ObjectHandlerBase> objectHandlers = BuildApplyHandlers(applyingInformation, calledFromHierarchy);

                    int count = objectHandlers.Count(o => o.ReportProgress && o.WillProvision(context, template, configuration)) + 1;
                    progressDelegate?.Invoke("Initializing engine", 1, count); // handlers + initializing message

                    tokenParser ??= await TokenParser.CreateAsync(context, template, applyingInformation).ConfigureAwait(false);

                    int step = 2;

                    // Remove artefacts a NoScript site would reject
                    var cleaner = new NoScriptTemplateCleaner(context);
                    if (messagesDelegate != null)
                    {
                        cleaner.MessagesDelegate = messagesDelegate;
                    }
                    template = await cleaner.CleanUpBeforeProvisioningAsync(template).ConfigureAwait(false);

                    await CallWebHooksAsync(template, tokenParser, ProvisioningTemplateWebhookKind.ProvisioningTemplateStarted).ConfigureAwait(false);

                    foreach (ObjectHandlerBase handler in objectHandlers)
                    {
                        if (!handler.WillProvision(context, template, configuration))
                        {
                            continue;
                        }

                        if (messagesDelegate != null)
                        {
                            handler.MessagesDelegate = messagesDelegate;
                        }

                        if (handler.ReportProgress && progressDelegate != null)
                        {
                            progressDelegate(handler.Name, step, count);
                            step++;
                        }

                        await CallWebHooksAsync(template, tokenParser,
                            ProvisioningTemplateWebhookKind.ObjectHandlerProvisioningStarted, handler.InternalName).ConfigureAwait(false);

                        try
                        {
                            // The parser is threaded, not copied: whatever this handler registered
                            // is visible to every handler after it. Do not "simplify" this to
                            // discard the return value.
                            tokenParser = await handler.ProvisionObjectsAsync(context, template, tokenParser, configuration).ConfigureAwait(false);
                        }
                        catch (Exception ex)
                        {
                            await CallWebHooksAsync(template, tokenParser,
                                ProvisioningTemplateWebhookKind.ExceptionOccurred, handler.InternalName, ex).ConfigureAwait(false);
                            throw;
                        }

                        await CallWebHooksAsync(template, tokenParser,
                            ProvisioningTemplateWebhookKind.ObjectHandlerProvisioningCompleted, handler.InternalName).ConfigureAwait(false);
                    }

                    // Notify the completed provisioning of the site
                    siteProvisionedDelegate?.Invoke(web.Title, web.Url.ToString());

                    await CallWebHooksAsync(template, tokenParser, ProvisioningTemplateWebhookKind.ProvisioningTemplateCompleted).ConfigureAwait(false);

                    return tokenParser;
                }
                finally
                {
                    System.Threading.Thread.CurrentThread.CurrentCulture = currentCultureInfoValue;
                }
            }
        }

        /// <summary>
        /// Builds the handler list for an apply run, in execution order.
        /// </summary>
        /// <remarks>
        /// <para><b>The three-pass ordering is load-bearing.</b> <c>ObjectField</c>,
        /// <c>ObjectContentType</c> and <c>ObjectListInstance</c> are each added three times, with
        /// <see cref="FieldAndListProvisioningStepHelper.Step"/> set to
        /// <c>ListAndStandardFields</c>, then <c>LookupFields</c>, then <c>ListSettings</c>.
        /// Without it, a lookup field pointing at a list defined later in the same template fails,
        /// because the list it targets does not exist yet. Any reordering here must be validated
        /// against test scenario 1.</para>
        /// <para>Handlers are registered as they are ported; the engine runs correctly with any
        /// subset, which is what makes phase-by-phase delivery possible.</para>
        /// </remarks>
        private static List<ObjectHandlerBase> BuildApplyHandlers(ProvisioningTemplateApplyingInformation applyingInformation, bool calledFromHierarchy)
        {
            var objectHandlers = new List<ObjectHandlerBase>();

            // MIGRATION PHASES 5-8: each handler is added here as it lands. The list below mirrors
            // PnP Framework's ApplyRemoteTemplate ordering exactly, including the three passes over
            // fields/content types/lists, so porting a handler is a one line change rather than a
            // sequencing decision.
            //
            //  Phase 5 (wave 1): RegionalSettings, SupportedUILanguages, SiteSettings, WebSettings,
            //                    Features, PropertyBagEntry, CustomActions, Theme, SiteHeader,
            //                    SiteFooter, ClientSidePages, SyntexModels, SearchSettings,
            //                    PersistTemplateInfo, ExtensibilityHandlers
            //  Phase 6 (wave 2): Localization, Field x3, ContentType x2, ListInstance x3,
            //                    ListInstanceDataRows, Files, SiteSecurity, TermGroups
            //  Phase 7 (wave 3): AuditSettings, SitePolicy, Workflows, Pages, PageContents,
            //                    Publishing, ComposedLook, ImageRenditions, Navigation
            //  Phase 8:          Tenant, ApplicationLifecycleManagement
            _ = applyingInformation;
            _ = calledFromHierarchy;

            return objectHandlers;
        }

        #endregion

        #region Extract a template

        /// <inheritdoc/>
        public async Task<ProvisioningTemplate> GetTemplateAsync(ExtractConfiguration configuration = null)
        {
            ILogger logger = context.Logger;
            using (logger?.BeginScope(PnPCoreProvisioningResources.Provisioning_ObjectHandlers_Extraction))
            {
                configuration ??= new ExtractConfiguration();

                ProvisioningTemplateCreationInformation creationInfo = configuration.ToCreationInformation();

                if (creationInfo.BaseTemplate != null)
                {
                    logger?.LogDebug(PnPCoreProvisioningResources.SiteToTemplateConversion_Base_template_available___0_, creationInfo.BaseTemplate.Id);
                }

                ProvisioningProgressDelegate progressDelegate = configuration.ProgressDelegate;
                ProvisioningMessagesDelegate messagesDelegate = configuration.MessagesDelegate;

                // Create empty object, hooking up the connector so the resulting template can be
                // applied to another site straight away
                var template = new ProvisioningTemplate
                {
                    Connector = configuration.FileConnector
                };

                List<ObjectHandlerBase> objectHandlers = BuildExtractHandlers(configuration);

                await context.Web.LoadAsync(w => w.Url).ConfigureAwait(false);

                int step = 1;
                int count = objectHandlers.Count(o => o.ReportProgress && o.WillExtract(context, template, configuration));

                foreach (ObjectHandlerBase handler in objectHandlers)
                {
                    if (!handler.WillExtract(context, template, configuration))
                    {
                        continue;
                    }

                    if (messagesDelegate != null)
                    {
                        handler.MessagesDelegate = messagesDelegate;
                    }

                    if (handler.ReportProgress && progressDelegate != null)
                    {
                        progressDelegate(handler.Name, step, count);
                        step++;
                    }

                    // PnP Framework cloned the ClientContext per handler so that one handler's
                    // pending queries could not disturb another's. PnP Core executes each call on
                    // its own, so the same context is used throughout - one fewer authentication
                    // round trip per handler.
                    template = await handler.ExtractObjectsAsync(context, template, configuration).ConfigureAwait(false);
                }

                return template;
            }
        }

        /// <inheritdoc/>
        public ProvisioningTemplate GetTemplate(ExtractConfiguration configuration = null)
        {
            return GetTemplateAsync(configuration).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Builds the handler list for an extract run, in execution order.
        /// </summary>
        private static List<ObjectHandlerBase> BuildExtractHandlers(ExtractConfiguration configuration)
        {
            var objectHandlers = new List<ObjectHandlerBase>();

            // MIGRATION PHASES 5-8: as with BuildApplyHandlers, handlers are registered here as
            // they land. Extraction uses a single pass - the Step.Export variant of the field,
            // content type and list handlers - rather than the three passes apply needs.
            _ = configuration;

            return objectHandlers;
        }

        #endregion

        #region Apply and extract a tenant template (hierarchy)

        /// <inheritdoc/>
        public async Task ApplyTenantTemplateAsync(ProvisioningHierarchy hierarchy, string sequenceId, ApplyConfiguration configuration = null)
        {
            if (hierarchy == null)
            {
                throw new ArgumentNullException(nameof(hierarchy));
            }

            ILogger logger = context.Logger;
            using (logger?.BeginScope(PnPCoreProvisioningResources.Provisioning_ObjectHandlers_Provisioning))
            {
                ProvisioningProgressDelegate progressDelegate = null;
                ProvisioningMessagesDelegate messagesDelegate = null;

                if (configuration == null)
                {
                    configuration = new ApplyConfiguration();
                }
                else
                {
                    // Apply the configuration parameters to the hierarchy
                    ApplyConfigurationParameters(configuration.Parameters, hierarchy.Parameters);

                    progressDelegate = configuration.ProgressDelegate;
                    messagesDelegate = configuration.MessagesDelegate;
                }

                logger?.LogDebug("{Source}: Attaching object handlers", Constants.LOGGING_SOURCE);

                List<ObjectHierarchyHandlerBase> objectHandlers = BuildHierarchyApplyHandlers();

                int count = objectHandlers.Count(o => o.ReportProgress && o.WillProvision(context, hierarchy, sequenceId, configuration)) + 1;
                progressDelegate?.Invoke("Initializing engine", 1, count); // handlers + initializing message

                int step = 2;

                TokenParser sequenceTokenParser = await TokenParser.CreateAsync(context, hierarchy).ConfigureAwait(false);

                await CallWebHooksAsync(hierarchy.Templates.FirstOrDefault(), sequenceTokenParser,
                    ProvisioningTemplateWebhookKind.ProvisioningStarted).ConfigureAwait(false);

                foreach (ObjectHierarchyHandlerBase handler in objectHandlers)
                {
                    if (!handler.WillProvision(context, hierarchy, sequenceId, configuration))
                    {
                        continue;
                    }

                    if (messagesDelegate != null)
                    {
                        handler.MessagesDelegate = messagesDelegate;
                    }

                    if (handler.ReportProgress && progressDelegate != null)
                    {
                        progressDelegate(handler.Name, step, count);
                        step++;
                    }

                    try
                    {
                        sequenceTokenParser = await handler.ProvisionObjectsAsync(context, hierarchy, sequenceId, sequenceTokenParser, configuration).ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        await CallWebHooksAsync(hierarchy.Templates.FirstOrDefault(), sequenceTokenParser,
                            ProvisioningTemplateWebhookKind.ProvisioningExceptionOccurred, handler.Name, ex).ConfigureAwait(false);
                        throw;
                    }
                }

                await CallWebHooksAsync(hierarchy.Templates.FirstOrDefault(), sequenceTokenParser,
                    ProvisioningTemplateWebhookKind.ProvisioningCompleted).ConfigureAwait(false);
            }
        }

        /// <inheritdoc/>
        public void ApplyTenantTemplate(ProvisioningHierarchy hierarchy, string sequenceId, ApplyConfiguration configuration = null)
        {
            ApplyTenantTemplateAsync(hierarchy, sequenceId, configuration).GetAwaiter().GetResult();
        }

        /// <inheritdoc/>
        public async Task<ProvisioningHierarchy> GetTenantTemplateAsync(ExtractConfiguration configuration = null)
        {
            configuration ??= new ExtractConfiguration();

            ILogger logger = context.Logger;
            using (logger?.BeginScope(PnPCoreProvisioningResources.Provisioning_ObjectHandlers_Extraction))
            {
                var tenantTemplate = new ProvisioningHierarchy
                {
                    Connector = configuration.FileConnector
                };

                List<ObjectHierarchyHandlerBase> objectHandlers = BuildHierarchyExtractHandlers(configuration);

                int step = 1;
                int count = objectHandlers.Count(o => o.ReportProgress && o.WillExtract(context, tenantTemplate, null, configuration));

                foreach (ObjectHierarchyHandlerBase handler in objectHandlers)
                {
                    if (!handler.WillExtract(context, tenantTemplate, null, configuration))
                    {
                        continue;
                    }

                    if (configuration.MessagesDelegate != null)
                    {
                        handler.MessagesDelegate = (message, type) => configuration.MessagesDelegate(message, type);
                    }

                    if (handler.ReportProgress && configuration.ProgressDelegate != null)
                    {
                        configuration.ProgressDelegate(handler.Name, step, count);
                        step++;
                    }

                    tenantTemplate = await handler.ExtractObjectsAsync(context, tenantTemplate, configuration).ConfigureAwait(false);
                }

                return tenantTemplate;
            }
        }

        /// <inheritdoc/>
        public ProvisioningHierarchy GetTenantTemplate(ExtractConfiguration configuration = null)
        {
            return GetTenantTemplateAsync(configuration).GetAwaiter().GetResult();
        }

        private static List<ObjectHierarchyHandlerBase> BuildHierarchyApplyHandlers()
        {
            // MIGRATION PHASE 8: ObjectHierarchyTenant, ObjectHierarchySequenceTermGroups,
            // ObjectHierarchySequenceSites, ObjectTeams, ObjectAzureActiveDirectory - in that order.
            return new List<ObjectHierarchyHandlerBase>();
        }

        private static List<ObjectHierarchyHandlerBase> BuildHierarchyExtractHandlers(ExtractConfiguration configuration)
        {
            // MIGRATION PHASE 8: ObjectHierarchySequenceSites when configuration.Tenant.Sequence is
            // set, ObjectTeams when configuration.Tenant.Teams is set.
            _ = configuration;
            return new List<ObjectHierarchyHandlerBase>();
        }

        #endregion

        #region Helpers

        private static void ApplyConfigurationParameters(Dictionary<string, string> sourceParameters, Dictionary<string, string> destParameters)
        {
            if (sourceParameters == null)
            {
                return;
            }

            foreach (KeyValuePair<string, string> p in sourceParameters)
            {
                destParameters[p.Key] = p.Value;
            }
        }

        private async Task WarnOnAsymmetricBaseTemplatesAsync(IWeb web, ProvisioningTemplate template, ProvisioningMessagesDelegate messagesDelegate)
        {
            if (string.IsNullOrEmpty(template.BaseSiteTemplate))
            {
                return;
            }

            await web.LoadAsync(w => w.WebTemplate, w => w.WebTemplateConfiguration).ConfigureAwait(false);
            string targetSiteTemplateId = $"{web.WebTemplate}#{BaseTemplates.BaseTemplateManager.GetConfiguration(web.WebTemplateConfiguration)}";

            if (targetSiteTemplateId.Equals(template.BaseSiteTemplate, StringComparison.InvariantCultureIgnoreCase))
            {
                return;
            }

            string warning = string.Format(CultureInfo.CurrentCulture,
                PnPCoreProvisioningResources.Provisioning_Asymmetric_Base_Templates,
                template.BaseSiteTemplate, targetSiteTemplateId);

            context.Logger?.LogWarning("{Source}: {Message}", Constants.LOGGING_SOURCE, warning);
            messagesDelegate?.Invoke(warning, ProvisioningMessageType.Warning);
        }

        private async Task CallWebHooksAsync(ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateWebhookKind kind,
            string objectHandler = null, Exception exception = null)
        {
            if (template == null)
            {
                return;
            }

            var webhooks = new List<ProvisioningWebhookBase>();

            // Merge the webhooks at template level with those at global level
            if (template.ProvisioningTemplateWebhooks != null && template.ProvisioningTemplateWebhooks.Any())
            {
                webhooks.AddRange(template.ProvisioningTemplateWebhooks);
            }

            if (template.ParentHierarchy?.ProvisioningWebhooks != null && template.ParentHierarchy.ProvisioningWebhooks.Any())
            {
                webhooks.AddRange(template.ParentHierarchy.ProvisioningWebhooks);
            }

            if (webhooks.Count == 0)
            {
                return;
            }

            // A plain client, not PnP Core's SharePoint one - see WebhookHttpClient for why.
            HttpClient httpClient = WebhookHttpClient.Instance;

            foreach (ProvisioningWebhookBase webhook in webhooks.Where(w => w.Kind == kind))
            {
                await WebhookSender.InvokeWebhookAsync(webhook, httpClient, kind, parser, objectHandler, exception, context.Logger).ConfigureAwait(false);
            }
        }

        #endregion
    }
}
