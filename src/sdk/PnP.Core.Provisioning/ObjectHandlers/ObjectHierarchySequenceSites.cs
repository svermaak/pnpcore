using Microsoft.Extensions.Logging;
using PnP.Core.Admin.Model.SharePoint;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.Provisioning.ObjectHandlers.TokenDefinitions;
using PnP.Core.Provisioning.ObjectHandlers.Utilities;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ClassicSiteCollectionModel = PnP.Core.Provisioning.Model.ClassicSiteCollection;
using CommunicationSiteCollectionModel = PnP.Core.Provisioning.Model.CommunicationSiteCollection;
using SiteCollectionModel = PnP.Core.Provisioning.Model.SiteCollection;
using TeamNoGroupSiteCollectionModel = PnP.Core.Provisioning.Model.TeamNoGroupSiteCollection;
using TeamSiteCollectionModel = PnP.Core.Provisioning.Model.TeamSiteCollection;
using TimeZone = PnP.Core.Admin.Model.SharePoint.TimeZone;

namespace PnP.Core.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Creates the site collections a tenant template's sequence declares, and applies the templates
    /// each of them names.
    /// </summary>
    /// <remarks>
    /// <para>Backlog <b>T14</b>, phase 8. This is the handler that turns a <c>&lt;pnp:Sequence&gt;</c>
    /// into real sites - four kinds of them - and then hands each new site to the ordinary template
    /// engine.</para>
    /// <para><b>Far smaller than the 918 lines it replaces.</b> PnP Framework carries its own site
    /// creation stack (<c>PnP.Framework.Sites.SiteCollection</c>, 1,589 lines) because CSOM's tenant
    /// API could not create modern sites; <c>ISiteCollectionManager</c> can, and it already handles
    /// the waiting, retrying and provisioning-status polling that most of those lines were.</para>
    /// <para><b>Existing sites are reused, not recreated.</b> A sequence is routinely re-applied, and
    /// the alternative to reuse is failing on the second run. Group connected sites are matched by
    /// alias and the rest by url, because that is what identifies each kind.</para>
    /// <para><b>Every site publishes its tokens before its templates are applied.</b> A template
    /// attached to one site in the sequence can refer to another by
    /// <c>{sequencesiteurl:…}</c> - which is the whole reason the sequence exists rather than being
    /// four separate applies - so all sites are created first and templates applied afterwards.</para>
    /// </remarks>
    internal class ObjectHierarchySequenceSites : ObjectHierarchyHandlerBase
    {
        public override string Name => "Sequences";

        public override bool WillProvision(PnPContext context, ProvisioningHierarchy hierarchy, string sequenceId,
            ApplyConfiguration configuration)
        {
            _willProvision ??= SequenceOf(hierarchy, sequenceId)?.SiteCollections?.Count > 0;
            return _willProvision.Value;
        }

        public override bool WillExtract(PnPContext context, ProvisioningHierarchy hierarchy, string sequenceId,
            ExtractConfiguration configuration)
        {
            // Extraction is phase 9 - see the remarks on ExtractObjectsAsync.
            _willExtract ??= false;
            return _willExtract.Value;
        }

        public override Task<ProvisioningHierarchy> ExtractObjectsAsync(PnPContext context, ProvisioningHierarchy hierarchy,
            ExtractConfiguration configuration)
        {
            // Reading a tenant back into a sequence needs the extract configuration's site url list
            // and hub site expansion, and it is only useful once the other hierarchy handlers can be
            // extracted alongside it. Deferred to phase 9 with ObjectFiles and ObjectPages, rather
            // than half-built here.
            return Task.FromResult(hierarchy);
        }

        #region Apply

        public override async Task<TokenParser> ProvisionObjectsAsync(PnPContext context, ProvisioningHierarchy hierarchy,
            string sequenceId, TokenParser parser, ApplyConfiguration configuration)
        {
            ProvisioningSequence sequence = SequenceOf(hierarchy, sequenceId);

            if (sequence == null || !(sequence.SiteCollections?.Count > 0))
            {
                return parser;
            }

            var created = new List<ProvisionedSite>();

            // Pass one: every site exists and every site's tokens are published.
            foreach (SiteCollectionModel siteCollection in sequence.SiteCollections)
            {
                ProvisionedSite site = await EnsureSiteAsync(context, siteCollection, parser, configuration)
                    .ConfigureAwait(false);

                if (site == null)
                {
                    continue;
                }

                created.Add(site);
                parser = await PublishTokensAsync(parser, siteCollection, site).ConfigureAwait(false);
            }

            // Pass two: hub registration, then templates. Hubs first so that a template applied to a
            // site can rely on its hub association already being in place.
            foreach (ProvisionedSite site in created)
            {
                await ApplyHubSettingsAsync(context, site, parser).ConfigureAwait(false);
            }

            foreach (ProvisionedSite site in created)
            {
                await ApplyTemplatesAsync(hierarchy, site, parser, configuration).ConfigureAwait(false);
            }

            foreach (ProvisionedSite site in created)
            {
                site.Context.Dispose();
            }

            return parser;
        }

        /// <summary>
        /// Creates a site collection, or opens the one that is already there.
        /// </summary>
        private async Task<ProvisionedSite> EnsureSiteAsync(PnPContext context, SiteCollectionModel siteCollection,
            TokenParser parser, ApplyConfiguration configuration)
        {
            string title = parser.ParseString(siteCollection.Title);

            try
            {
                PnPContext siteContext = await OpenExistingAsync(context, siteCollection, parser).ConfigureAwait(false);

                if (siteContext != null)
                {
                    WriteMessage($"Using existing site {siteContext.Uri}", ProvisioningMessageType.Progress);
                }
                else
                {
                    WriteMessage($"Creating site {title}", ProvisioningMessageType.Progress);

                    siteContext = await CreateAsync(context, siteCollection, parser, configuration).ConfigureAwait(false);

                    if (siteContext == null)
                    {
                        return null;
                    }

                    WriteMessage($"Created site {siteContext.Uri}", ProvisioningMessageType.Progress);
                }

                await siteContext.Site.LoadAsync(s => s.Id, s => s.GroupId).ConfigureAwait(false);

                return new ProvisionedSite
                {
                    Model = siteCollection,
                    Context = siteContext,
                };
            }
            catch (Exception ex)
            {
                string warning = $"The site '{title}' could not be provisioned, so it and the templates " +
                    $"attached to it were skipped: {ErrorText.Describe(ex)}";
                context.Logger?.LogWarning(ex, "{Source}: {Message}", Constants.LOGGING_SOURCE, warning);
                WriteMessage(warning, ProvisioningMessageType.Warning);
                return null;
            }
        }

        /// <summary>
        /// Opens the site if it already exists.
        /// </summary>
        /// <remarks>
        /// <b>Group connected sites are matched by alias, not by url.</b> A team site's url is
        /// derived from its alias by SharePoint and may differ from what the template guessed - and
        /// the alias is what actually collides, so a url check would report "not there", try to
        /// create, and fail on a duplicate alias.
        /// </remarks>
        private static async Task<PnPContext> OpenExistingAsync(PnPContext context, SiteCollectionModel siteCollection,
            TokenParser parser)
        {
            ISiteCollectionManager manager = context.GetSiteCollectionManager();

            if (siteCollection is TeamSiteCollectionModel team)
            {
                string alias = SanitizeAlias(parser.ParseString(team.Alias));

                ISiteCollectionWithDetails match = (await manager.GetSiteCollectionsWithDetailsAsync()
                    .ConfigureAwait(false))
                    .FirstOrDefault(s => string.Equals(AliasOf(s.Url), alias, StringComparison.OrdinalIgnoreCase));

                return match == null ? null : await context.CloneAsync(match.Url).ConfigureAwait(false);
            }

            Uri url = UrlOf(siteCollection, parser);

            if (url == null || !await manager.SiteExistsAsync(url).ConfigureAwait(false))
            {
                return null;
            }

            return await context.CloneAsync(url).ConfigureAwait(false);
        }

        private static async Task<PnPContext> CreateAsync(PnPContext context, SiteCollectionModel siteCollection,
            TokenParser parser, ApplyConfiguration configuration)
        {
            ISiteCollectionManager manager = context.GetSiteCollectionManager();

            var creationOptions = new SiteCreationOptions
            {
                UsingApplicationPermissions = false,

                // The template engine runs against the new site immediately afterwards, so waiting
                // for it to be fully provisioned is not optional here the way it is for a caller who
                // only wants the site to exist.
                WaitForAsyncProvisioning = true,
            };

            CommonSiteOptions options = BuildOptions(siteCollection, parser);

            return options == null
                ? null
                : await manager.CreateSiteCollectionAsync(options, creationOptions).ConfigureAwait(false);
        }

        /// <summary>
        /// Turns a template's site collection into the options PnP Core creates from.
        /// </summary>
        private static CommonSiteOptions BuildOptions(SiteCollectionModel siteCollection, TokenParser parser)
        {
            string title = parser.ParseString(siteCollection.Title);
            string description = parser.ParseString(siteCollection.Description);

            switch (siteCollection)
            {
                case TeamSiteCollectionModel team:
                {
                    var options = new TeamSiteOptions(SanitizeAlias(parser.ParseString(team.Alias)), title)
                    {
                        Description = description,
                        Classification = parser.ParseString(team.Classification),
                        IsPublic = team.IsPublic,
                        Language = LanguageOf(team.Language),
                    };

                    if (Guid.TryParse(parser.ParseString(team.SiteDesign), out Guid designId))
                    {
                        options.SiteDesignId = designId;
                    }

                    return options;
                }

                case CommunicationSiteCollectionModel communication:
                {
                    var options = new CommunicationSiteOptions(
                        new Uri(parser.ParseString(communication.Url)), title)
                    {
                        Description = description,
                        Classification = parser.ParseString(communication.Classification),
                        Owner = parser.ParseString(communication.Owner),
                        ShareByEmailEnabled = communication.AllowFileSharingForGuestUsers,
                        Language = LanguageOf(communication.Language),
                    };

                    // A site design is either one of the three built in designs, named, or a custom
                    // one addressed by id - the two go in different properties.
                    string design = parser.ParseString(communication.SiteDesign);

                    if (Guid.TryParse(design, out Guid customDesign))
                    {
                        options.SiteDesignId = customDesign;
                    }
                    else if (Enum.TryParse(design, true, out CommunicationSiteDesign builtIn))
                    {
                        options.SiteDesign = builtIn;
                    }

                    return options;
                }

                case TeamNoGroupSiteCollectionModel noGroup:
                {
                    return new TeamSiteWithoutGroupOptions(
                        new Uri(parser.ParseString(noGroup.Url)), title)
                    {
                        Description = description,
                        Classification = parser.ParseString(noGroup.Classification),
                        Owner = parser.ParseString(noGroup.Owner),
                        Language = LanguageOf(noGroup.Language),
                        TimeZone = TimeZoneOf(noGroup.TimeZoneId),
                    };
                }

                case ClassicSiteCollectionModel classic:
                {
                    return new ClassicSiteOptions(
                        new Uri(parser.ParseString(classic.Url)),
                        title,
                        parser.ParseString(classic.WebTemplate),
                        parser.ParseString(classic.Owner),
                        LanguageOf(classic.Language),
                        TimeZoneOf(classic.TimeZoneId) ?? TimeZone.UTCPLUS0100_BRUSSELS_COPENHAGEN_MADRID_PARIS);
                }

                default:
                    return null;
            }
        }

        #endregion

        #region Tokens

        /// <summary>
        /// Publishes the tokens other parts of the hierarchy address this site by.
        /// </summary>
        /// <remarks>
        /// Only for a site that declares a <c>ProvisioningID</c> - that attribute is what the tokens
        /// are keyed on, and a site without one cannot be referred to.
        /// </remarks>
        private static async Task<TokenParser> PublishTokensAsync(TokenParser parser, SiteCollectionModel siteCollection,
            ProvisionedSite site)
        {
            string provisioningId = siteCollection.ProvisioningId;

            if (string.IsNullOrEmpty(provisioningId))
            {
                return parser;
            }

            parser.AddToken(new SequenceSiteUrlUrlToken(site.Context, provisioningId, site.Context.Uri.ToString()));
            parser.AddToken(new SequenceSiteIdToken(site.Context, provisioningId, site.Context.Site.Id));
            parser.AddToken(new SequenceSiteCollectionIdToken(site.Context, provisioningId, site.Context.Site.Id));

            if (site.Context.Site.GroupId != Guid.Empty)
            {
                parser.AddToken(new SequenceSiteGroupIdToken(site.Context, provisioningId, site.Context.Site.GroupId));
            }

            await Task.CompletedTask.ConfigureAwait(false);

            return parser;
        }

        #endregion

        #region Hub sites

        /// <summary>
        /// Registers the site as a hub, or leaves it alone.
        /// </summary>
        /// <remarks>
        /// <b>Registering twice fails.</b> A site that is already a hub answers <c>GetHubSiteDataAsync</c>,
        /// which is cheaper than registering and catching - and re-applying a sequence is the normal
        /// case, not the exception.
        /// </remarks>
        private async Task ApplyHubSettingsAsync(PnPContext context, ProvisionedSite site, TokenParser parser)
        {
            if (!site.Model.IsHubSite)
            {
                return;
            }

            try
            {
                IHubSite hub = await ExistingHubAsync(site.Context).ConfigureAwait(false);

                if (hub == null)
                {
                    WriteMessage($"Registering {site.Context.Uri} as a hub site", ProvisioningMessageType.Progress);

                    hub = await site.Context.Site.RegisterHubSiteAsync().ConfigureAwait(false);
                }

                string title = parser.ParseString(site.Model.HubSiteTitle);
                string logo = parser.ParseString(site.Model.HubSiteLogoUrl);

                if (hub != null && (!string.IsNullOrEmpty(title) || !string.IsNullOrEmpty(logo)))
                {
                    if (!string.IsNullOrEmpty(title))
                    {
                        hub.Title = title;
                    }

                    if (!string.IsNullOrEmpty(logo))
                    {
                        hub.LogoUrl = logo;
                    }

                    await hub.UpdateAsync().ConfigureAwait(false);
                }
            }
            catch (Exception ex)
            {
                string warning = $"{site.Context.Uri} could not be registered as a hub site: {ErrorText.Describe(ex)}";
                context.Logger?.LogWarning(ex, "{Source}: {Message}", Constants.LOGGING_SOURCE, warning);
                WriteMessage(warning, ProvisioningMessageType.Warning);
            }
        }

        private static async Task<IHubSite> ExistingHubAsync(PnPContext siteContext)
        {
            try
            {
                await siteContext.Site.LoadAsync(s => s.IsHubSite).ConfigureAwait(false);

                return siteContext.Site.IsHubSite
                    ? await siteContext.Site.GetHubSiteDataAsync(null).ConfigureAwait(false)
                    : null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        #endregion

        #region Templates

        /// <summary>
        /// Applies the hierarchy's templates that this site names, in the order it names them.
        /// </summary>
        /// <remarks>
        /// <b>The token parser is threaded through, not rebuilt.</b> It carries the sequence tokens
        /// every site published, which is what lets one site's template refer to another site.
        /// </remarks>
        private async Task ApplyTemplatesAsync(ProvisioningHierarchy hierarchy, ProvisionedSite site,
            TokenParser parser, ApplyConfiguration configuration)
        {
            if (!(site.Model.Templates?.Count > 0))
            {
                return;
            }

            foreach (string templateId in site.Model.Templates)
            {
                string id = parser.ParseString(templateId);

                ProvisioningTemplate template = hierarchy.Templates.FirstOrDefault(
                    t => string.Equals(t.Id, id, StringComparison.OrdinalIgnoreCase));

                if (template == null)
                {
                    string warning = $"The hierarchy has no template with id '{id}', which " +
                        $"{site.Context.Uri} refers to, so it was skipped.";
                    site.Context.Logger?.LogWarning("{Source}: {Message}", Constants.LOGGING_SOURCE, warning);
                    WriteMessage(warning, ProvisioningMessageType.Warning);
                    continue;
                }

                try
                {
                    WriteMessage($"Applying template {id} to {site.Context.Uri}", ProvisioningMessageType.Progress);

                    // The connector belongs to the hierarchy, and a template inside one does not
                    // carry its own - without this every file and every site script in the template
                    // resolves to nothing.
                    template.Connector ??= hierarchy.Connector;

                    var manager = (ProvisioningManager)site.Context.GetProvisioningManager();

                    await manager.ApplyTemplateAsync(template, configuration,
                        calledFromHierarchy: true, parser).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    string warning = $"The template '{id}' could not be applied to {site.Context.Uri}: " +
                        ErrorText.Describe(ex);
                    site.Context.Logger?.LogWarning(ex, "{Source}: {Message}", Constants.LOGGING_SOURCE, warning);
                    WriteMessage(warning, ProvisioningMessageType.Warning);
                }
            }
        }

        #endregion

        #region Helpers

        private static ProvisioningSequence SequenceOf(ProvisioningHierarchy hierarchy, string sequenceId)
        {
            if (hierarchy?.Sequences == null)
            {
                return null;
            }

            return string.IsNullOrEmpty(sequenceId)
                ? hierarchy.Sequences.FirstOrDefault()
                : hierarchy.Sequences.FirstOrDefault(s => s.ID == sequenceId);
        }

        private static Uri UrlOf(SiteCollectionModel siteCollection, TokenParser parser)
        {
            string url = siteCollection switch
            {
                CommunicationSiteCollectionModel communication => communication.Url,
                TeamNoGroupSiteCollectionModel noGroup => noGroup.Url,
                ClassicSiteCollectionModel classic => classic.Url,
                _ => null,
            };

            url = parser.ParseString(url);

            return Uri.TryCreate(url, UriKind.Absolute, out Uri parsed) ? parsed : null;
        }

        /// <summary>
        /// The alias part of a group connected site's url.
        /// </summary>
        private static string AliasOf(Uri url)
        {
            string path = url?.AbsolutePath?.TrimEnd('/');

            return string.IsNullOrEmpty(path) ? null : path.Substring(path.LastIndexOf('/') + 1);
        }

        /// <summary>
        /// Removes what SharePoint will not accept in a group alias.
        /// </summary>
        /// <remarks>
        /// <para>PnP Framework strips disallowed characters and folds accents before creating. Left
        /// as is, a perfectly ordinary title-derived alias produces a creation failure whose message
        /// is about the alias rather than about the character in it.</para>
        /// <para>Fully qualified: <c>PnP.Core.Model.SharePoint</c> has a <c>UrlUtility</c> of its own,
        /// and the unqualified name binds to that one - which has neither method.</para>
        /// </remarks>
        private static string SanitizeAlias(string alias)
        {
            if (string.IsNullOrEmpty(alias))
            {
                return alias;
            }

            return Provisioning.Utilities.UrlUtility.RemoveUnallowedCharacters(
                Provisioning.Utilities.UrlUtility.ReplaceAccentedCharactersWithLatin(alias));
        }

        /// <summary>
        /// Maps an LCID to PnP Core's language, falling back to English.
        /// </summary>
        /// <remarks>
        /// A template may carry a language PnP Core's enum does not name, and a site in the wrong
        /// language is better than no site plus an unexplained failure - but it is said out loud.
        /// </remarks>
        private static Language LanguageOf(int lcid)
        {
            return Enum.IsDefined(typeof(Language), lcid) ? (Language)lcid : Language.Default;
        }

        private static TimeZone? TimeZoneOf(int id)
        {
            return Enum.IsDefined(typeof(TimeZone), id) ? (TimeZone)id : (TimeZone?)null;
        }

        /// <summary>
        /// A site the sequence created or adopted, and the context bound to it.
        /// </summary>
        private sealed class ProvisionedSite
        {
            internal SiteCollectionModel Model { get; set; }

            internal PnPContext Context { get; set; }
        }

        #endregion
    }
}
