using Microsoft.Extensions.Logging;
using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using CorePermissions = PnP.Core.Model.SharePoint.BasePermissions;
using CustomActionModel = PnP.Core.Provisioning.Model.CustomAction;
using TemplatePermissions = PnP.Core.Provisioning.Model.BasePermissions;

namespace PnP.Core.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Provisions and extracts user custom actions at both site collection and web scope.
    /// </summary>
    /// <remarks>
    /// <para><b>Scope matters.</b> Site collection scoped actions are only touched when the target
    /// is a root web - applying them from a sub site would write to the whole site collection, which
    /// a sub site template has no business doing. Extraction follows the same rule, so a sub site
    /// template does not inherit its parent's actions.</para>
    /// <para><b>NoScript sites</b> reject custom actions that carry script, which is most of them.
    /// One exception survives: an action whose <c>ClientSideComponentId</c> is set is an SPFx
    /// extension, not script, and is allowed. Anything else is skipped with a warning rather than
    /// left to fail mid-run.</para>
    /// <para><b>Not yet ported:</b> <c>{res:}</c> localization of a custom action's title and
    /// description. The CSOM plumbing exists (see <c>UserResourcePath.ForUserCustomAction</c>) but
    /// the resource-token machinery belongs to <c>ObjectLocalization</c> in phase 6. A template that
    /// uses it gets a warning here - the literal token would otherwise be written as the title.</para>
    /// </remarks>
    internal class ObjectCustomActions : ObjectHandlerBase
    {
        public override string Name => "Custom Actions";

        public override string InternalName => "CustomActions";

        public override bool WillProvision(PnPContext context, ProvisioningTemplate template, ApplyConfiguration configuration)
        {
            _willProvision ??= template.CustomActions != null
                && (template.CustomActions.SiteCustomActions.Any() || template.CustomActions.WebCustomActions.Any());

            return _willProvision.Value;
        }

        public override bool WillExtract(PnPContext context, ProvisioningTemplate template, ExtractConfiguration configuration)
        {
            // PnP Framework queried the site here to decide. That is I/O in a synchronous predicate
            // the engine calls twice per run; extraction of an empty collection costs one request
            // and produces the same template.
            return true;
        }

        #region Apply

        public override async Task<TokenParser> ProvisionObjectsAsync(PnPContext context, ProvisioningTemplate template, TokenParser parser, ApplyConfiguration configuration)
        {
            using (context.Logger?.BeginScope(Name))
            {
                if (template.CustomActions == null)
                {
                    return parser;
                }

                IWeb web = await context.Web.GetAsync(w => w.ServerRelativeUrl, w => w.UserCustomActions).ConfigureAwait(false);
                await context.Site.LoadAsync(s => s.ServerRelativeUrl, s => s.UserCustomActions).ConfigureAwait(false);

                bool isNoScriptSite = await context.Web.IsNoScriptSiteAsync().ConfigureAwait(false);

                // A sub site must not write site collection scoped actions.
                //
                // PnP Framework rebased the token parser onto the root web before doing so and back
                // afterwards. That is unnecessary here: the branch only runs when this web IS the
                // root web, so the rebase was always a no-op.
                if (!IsSubSite(web))
                {
                    await ApplyScopeAsync(context, context.Site.UserCustomActions,
                        template.CustomActions.SiteCustomActions, parser, isNoScriptSite, "site").ConfigureAwait(false);
                }

                await ApplyScopeAsync(context, context.Web.UserCustomActions,
                    template.CustomActions.WebCustomActions, parser, isNoScriptSite, "web").ConfigureAwait(false);

                return parser;
            }
        }

        private async Task ApplyScopeAsync(PnPContext context, IUserCustomActionCollection collection,
            IEnumerable<CustomActionModel> customActions, TokenParser parser, bool isNoScriptSite, string scope)
        {
            List<CustomActionModel> wanted = customActions?.ToList() ?? new List<CustomActionModel>();
            if (wanted.Count == 0)
            {
                return;
            }

            // AsRequested rather than ToListAsync: the collection was loaded with its parent above,
            // and .NET 10 makes an unqualified ToListAsync ambiguous between PnP Core's queryable
            // extension and System.Linq.AsyncEnumerable.
            List<IUserCustomAction> existingActions = collection.AsRequested().ToList();

            foreach (CustomActionModel customAction in wanted)
            {
                IUserCustomAction existing = existingActions
                    .FirstOrDefault(a => string.Equals(a.Name, customAction.Name, StringComparison.Ordinal));

                if (existing != null && customAction.Remove)
                {
                    await existing.DeleteAsync().ConfigureAwait(false);
                    existingActions.Remove(existing);
                    continue;
                }

                // Only script-carrying actions are refused on a NoScript site; an SPFx extension is
                // identified by its client side component id and is allowed.
                if (isNoScriptSite && customAction.ClientSideComponentId == Guid.Empty)
                {
                    string message = $"This is a NoScript site, so the custom action '{customAction.Name}' was skipped.";
                    context.Logger?.LogWarning("{Source}: {Message}", Constants.LOGGING_SOURCE, message);
                    WriteMessage(message, ProvisioningMessageType.Warning);
                    continue;
                }

                WarnOnUnsupportedResourceTokens(context, customAction);

                if (existing == null)
                {
                    if (customAction.Remove || !customAction.Enabled)
                    {
                        continue;
                    }

                    context.Logger?.LogDebug("{Source}: adding {Scope} scoped custom action '{Name}'.",
                        Constants.LOGGING_SOURCE, scope, customAction.Name);

                    existingActions.Add(await collection.AddAsync(BuildOptions(customAction, parser)).ConfigureAwait(false));
                }
                else
                {
                    await UpdateAsync(context, existing, customAction, parser).ConfigureAwait(false);
                }
            }
        }

        private static AddUserCustomActionOptions BuildOptions(CustomActionModel customAction, TokenParser parser)
        {
            return new AddUserCustomActionOptions
            {
                ClientSideComponentId = customAction.ClientSideComponentId != Guid.Empty ? customAction.ClientSideComponentId : null,
                ClientSideComponentProperties = parser.ParseString(customAction.ClientSideComponentProperties),
                CommandUIExtension = customAction.CommandUIExtension != null
                    ? parser.ParseString(customAction.CommandUIExtension.ToString())
                    : string.Empty,
                Description = parser.ParseString(customAction.Description),
                Group = customAction.Group,
                ImageUrl = parser.ParseString(customAction.ImageUrl),
                Location = customAction.Location,
                Name = customAction.Name,
                RegistrationId = parser.ParseString(customAction.RegistrationId),
                RegistrationType = customAction.RegistrationType,
                Rights = ToCorePermissions(customAction.Rights),
                ScriptBlock = parser.ParseString(customAction.ScriptBlock),
                ScriptSrc = parser.ParseString(customAction.ScriptSrc),
                Sequence = customAction.Sequence,
                Title = parser.ParseString(customAction.Title),
                Url = parser.ParseString(customAction.Url),
            };
        }

        /// <summary>
        /// Brings an existing custom action in line with the template, updating only when something
        /// actually differs.
        /// </summary>
        private static async Task UpdateAsync(PnPContext context, IUserCustomAction existing, CustomActionModel customAction, TokenParser parser)
        {
            bool dirty = false;

            // A null CommandUIExtension in the template means "blank it out", which is how a delta
            // template removes a ribbon customization. Treating null as "leave alone" would make
            // that impossible to express.
            string commandUIExtension = customAction.CommandUIExtension != null
                ? parser.ParseString(customAction.CommandUIExtension.ToString())
                : null;

            dirty |= SetIfChanged(existing.CommandUIExtension, commandUIExtension, v => existing.CommandUIExtension = v);

            if (customAction.ClientSideComponentId != Guid.Empty)
            {
                dirty |= SetIfChanged(existing.ClientSideComponentId, customAction.ClientSideComponentId, v => existing.ClientSideComponentId = v);
            }

            if (!string.IsNullOrEmpty(customAction.ClientSideComponentProperties))
            {
                dirty |= SetIfChanged(existing.ClientSideComponentProperties,
                    parser.ParseString(customAction.ClientSideComponentProperties), v => existing.ClientSideComponentProperties = v);
            }

            dirty |= SetIfChanged(existing.Description, customAction.Description, v => existing.Description = v);
            dirty |= SetIfChanged(existing.Group, customAction.Group, v => existing.Group = v);
            dirty |= SetIfChanged(existing.ImageUrl, parser.ParseString(customAction.ImageUrl), v => existing.ImageUrl = v);
            dirty |= SetIfChanged(existing.Location, customAction.Location, v => existing.Location = v);
            dirty |= SetIfChanged(existing.RegistrationId, parser.ParseString(customAction.RegistrationId), v => existing.RegistrationId = v);
            dirty |= SetIfChanged(existing.RegistrationType, customAction.RegistrationType, v => existing.RegistrationType = v);
            dirty |= SetIfChanged(existing.ScriptBlock, parser.ParseString(customAction.ScriptBlock), v => existing.ScriptBlock = v);
            dirty |= SetIfChanged(existing.ScriptSrc, parser.ParseString(customAction.ScriptSrc), v => existing.ScriptSrc = v);
            dirty |= SetIfChanged(existing.Sequence, customAction.Sequence, v => existing.Sequence = v);
            dirty |= SetIfChanged(existing.Title, parser.ParseString(customAction.Title), v => existing.Title = v);
            dirty |= SetIfChanged(existing.Url, parser.ParseString(customAction.Url), v => existing.Url = v);

            if (dirty)
            {
                context.Logger?.LogDebug("{Source}: updating custom action '{Name}'.", Constants.LOGGING_SOURCE, customAction.Name);
                await existing.UpdateAsync().ConfigureAwait(false);
            }
        }

        private void WarnOnUnsupportedResourceTokens(PnPContext context, CustomActionModel customAction)
        {
            if (!ContainsResourceToken(customAction.Title) && !ContainsResourceToken(customAction.Description))
            {
                return;
            }

            string message = string.Format(System.Globalization.CultureInfo.CurrentCulture,
                "The custom action '{0}' uses a {{res:...}} token. Per-language values are written by " +
                "ObjectLocalization, which lands in phase 6 - only the default language value was applied.",
                customAction.Name);

            context.Logger?.LogWarning("{Source}: {Message}", Constants.LOGGING_SOURCE, message);
            WriteMessage(message, ProvisioningMessageType.Warning);
        }

        #endregion

        #region Extract

        public override async Task<ProvisioningTemplate> ExtractObjectsAsync(PnPContext context, ProvisioningTemplate template, ExtractConfiguration configuration)
        {
            using (context.Logger?.BeginScope(Name))
            {
                IWeb web = await context.Web.GetAsync(w => w.ServerRelativeUrl, w => w.UserCustomActions).ConfigureAwait(false);
                await context.Site.LoadAsync(s => s.ServerRelativeUrl, s => s.UserCustomActions).ConfigureAwait(false);

                bool isSubSite = IsSubSite(web);

                var customActions = new CustomActions();

                foreach (IUserCustomAction customAction in web.UserCustomActions.AsRequested())
                {
                    customActions.WebCustomActions.Add(Copy(customAction));
                }

                // As on apply: a sub site's template does not carry the site collection's actions.
                if (!isSubSite)
                {
                    foreach (IUserCustomAction customAction in context.Site.UserCustomActions.AsRequested())
                    {
                        customActions.SiteCustomActions.Add(Copy(customAction));
                    }
                }

                template.CustomActions = customActions;

                ProvisioningTemplate baseTemplate = configuration?.ToCreationInformation()?.BaseTemplate;
                if (baseTemplate != null)
                {
                    RemoveBaseTemplateEntries(context, template, baseTemplate, isSubSite);
                }

                return template;
            }
        }

        /// <summary>
        /// Drops the actions the target site template already provides, so the extracted template
        /// carries only what a person actually added.
        /// </summary>
        private static void RemoveBaseTemplateEntries(PnPContext context, ProvisioningTemplate template, ProvisioningTemplate baseTemplate, bool isSubSite)
        {
            if (!isSubSite)
            {
                RemoveMatching(context, template.CustomActions.SiteCustomActions, baseTemplate.CustomActions.SiteCustomActions, "site");
            }

            RemoveMatching(context, template.CustomActions.WebCustomActions, baseTemplate.CustomActions.WebCustomActions, "web");
        }

        private static void RemoveMatching(PnPContext context, CustomActionCollection target, CustomActionCollection baseline, string scope)
        {
            foreach (CustomActionModel customAction in baseline)
            {
                int index = target.FindIndex(c => c.Name.Equals(customAction.Name, StringComparison.Ordinal));
                if (index < 0)
                {
                    continue;
                }

                context.Logger?.LogDebug("{Source}: dropping {Scope} scoped custom action '{Name}' - it is part of the base template.",
                    Constants.LOGGING_SOURCE, scope, customAction.Name);

                target.RemoveAt(index);
            }
        }

        private static CustomActionModel Copy(IUserCustomAction userCustomAction)
        {
            return new CustomActionModel
            {
                Description = userCustomAction.Description,
                Enabled = true,
                Group = userCustomAction.Group,
                ImageUrl = userCustomAction.ImageUrl,
                Location = userCustomAction.Location,
                Name = userCustomAction.Name,
                Rights = ToTemplatePermissions(userCustomAction.Rights),
                ScriptBlock = userCustomAction.ScriptBlock,
                ScriptSrc = userCustomAction.ScriptSrc,
                Sequence = userCustomAction.Sequence,
                Title = userCustomAction.Title,
                Url = userCustomAction.Url,
                RegistrationId = userCustomAction.RegistrationId,
                RegistrationType = userCustomAction.RegistrationType,
                ClientSideComponentId = userCustomAction.ClientSideComponentId,
                ClientSideComponentProperties = userCustomAction.ClientSideComponentProperties,
                CommandUIExtension = !string.IsNullOrEmpty(userCustomAction.CommandUIExtension)
                    ? XElement.Parse(userCustomAction.CommandUIExtension)
                    : null,
            };

            // MIGRATION PHASE 6: PersistMultiLanguageResources also wrote the title and description
            // into the template's resource files here and replaced them with {res:} tokens. That
            // needs the resource file plumbing ObjectLocalization owns.
        }

        #endregion

        #region Helpers

        private static bool SetIfChanged<T>(T current, T wanted, Action<T> set)
        {
            if (EqualityComparer<T>.Default.Equals(current, wanted))
            {
                return false;
            }

            set(wanted);
            return true;
        }

        private static bool ContainsResourceToken(string value)
        {
            return !string.IsNullOrEmpty(value)
                && value.IndexOf("{res:", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>
        /// Converts the template's plain permission mask into the live model PnP Core expects.
        /// </summary>
        /// <remarks>
        /// The provisioning model deliberately keeps its own <c>BasePermissions</c> POCO so a
        /// template can round trip with no connection - see the remarks on that class. This is the
        /// one place the two meet.
        /// </remarks>
        private static IBasePermissions ToCorePermissions(TemplatePermissions permissions)
        {
            var rights = new CorePermissions();

            if (permissions != null)
            {
                rights.Low = permissions.Low;
                rights.High = permissions.High;
            }

            return rights;
        }

        private static TemplatePermissions ToTemplatePermissions(IBasePermissions permissions)
        {
            if (permissions == null)
            {
                return new TemplatePermissions();
            }

            return new TemplatePermissions
            {
                Low = (uint)permissions.Low,
                High = (uint)permissions.High,
            };
        }

        #endregion
    }
}
