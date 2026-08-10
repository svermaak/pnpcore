using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.Services;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Applies a tenant template's <c>&lt;pnp:Tenant&gt;</c> element, once, before any site in the
    /// sequence is created.
    /// </summary>
    /// <remarks>
    /// <para>Phase 8. <b>The work is <see cref="ObjectTenant"/>'s</b> - the element is the same one,
    /// and PnP Framework's two handlers call an identical list of eleven helpers. Only two things
    /// differ, and both are inputs rather than behaviour: the element hangs off the hierarchy, and
    /// the files come from the hierarchy's connector.</para>
    /// <para><b>Why it runs first.</b> A sequence's sites are routinely created from a site design
    /// the same template defines, so the designs, scripts and storage entities have to exist - and
    /// have published their tokens - before <c>ObjectHierarchySequenceSites</c> runs. That ordering
    /// is why <c>ObjectTenant</c> is skipped when a template is applied from within a hierarchy:
    /// otherwise every tenant-wide write would repeat once per site.</para>
    /// </remarks>
    internal class ObjectHierarchyTenant : ObjectHierarchyHandlerBase
    {
        public override string Name => "Tenant Settings";

        public override bool WillProvision(PnPContext context, ProvisioningHierarchy hierarchy, string sequenceId,
            ApplyConfiguration configuration)
        {
            ProvisioningTenant tenant = hierarchy?.Tenant;

            _willProvision ??= tenant != null
                && (tenant.AppCatalog != null
                    || tenant.ContentDeliveryNetwork != null
                    || tenant.SiteDesigns?.Count > 0
                    || tenant.SiteScripts?.Count > 0
                    || tenant.StorageEntities?.Count > 0
                    || tenant.WebApiPermissions?.Count > 0
                    || tenant.Themes?.Count > 0
                    || tenant.SPUsersProfiles?.Count > 0
                    || tenant.Office365GroupLifecyclePolicies?.Count > 0
                    || tenant.Office365GroupsSettings?.Properties?.Count > 0
                    || tenant.SharingSettings != null);

            return _willProvision.Value;
        }

        public override bool WillExtract(PnPContext context, ProvisioningHierarchy hierarchy, string sequenceId,
            ExtractConfiguration configuration)
        {
            // Tenant settings are not extracted by default - matching ObjectTenant and PnP Framework.
            _willExtract ??= false;
            return _willExtract.Value;
        }

        public override Task<ProvisioningHierarchy> ExtractObjectsAsync(PnPContext context, ProvisioningHierarchy hierarchy,
            ExtractConfiguration configuration)
        {
            return Task.FromResult(hierarchy);
        }

        public override async Task<TokenParser> ProvisionObjectsAsync(PnPContext context, ProvisioningHierarchy hierarchy,
            string sequenceId, TokenParser parser, ApplyConfiguration configuration)
        {
            // The inner handler's messages are this handler's messages: the caller sees one
            // "Tenant Settings" step either way, and losing the delegate would silence every
            // progress line and every warning the tenant work produces.
            var inner = new ObjectTenant
            {
                MessagesDelegate = MessagesDelegate,
            };

            return await inner.ApplyTenantAsync(context, hierarchy?.Tenant, hierarchy?.Connector, parser)
                .ConfigureAwait(false);
        }
    }
}
