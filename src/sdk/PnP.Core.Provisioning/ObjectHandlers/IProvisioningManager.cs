using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Applies provisioning templates to a site, and extracts them back out of one.
    /// </summary>
    /// <remarks>
    /// Obtained from a context - <c>context.GetProvisioningManager()</c>. This replaces PnP
    /// Framework's five CSOM extension methods (<c>Web.ApplyProvisioningTemplate</c>,
    /// <c>Web.GetProvisioningTemplate</c>, <c>Tenant.ApplyTenantTemplate</c>,
    /// <c>Tenant.GetTenantTemplate</c>) with the manager-from-context shape the rest of the SDK
    /// uses.
    /// </remarks>
    public interface IProvisioningManager
    {
        /// <summary>
        /// Applies a provisioning template to the site this manager was obtained from.
        /// </summary>
        /// <param name="template">The template to apply</param>
        /// <param name="configuration">Which handlers to run and how, or <c>null</c> for all of them</param>
        Task ApplyTemplateAsync(ProvisioningTemplate template, ApplyConfiguration configuration = null);

        /// <summary>
        /// Applies a provisioning template to the site this manager was obtained from.
        /// </summary>
        /// <param name="template">The template to apply</param>
        /// <param name="configuration">Which handlers to run and how, or <c>null</c> for all of them</param>
        void ApplyTemplate(ProvisioningTemplate template, ApplyConfiguration configuration = null);

        /// <summary>
        /// Extracts a provisioning template from the site this manager was obtained from.
        /// </summary>
        /// <param name="configuration">What to extract and how, or <c>null</c> for everything</param>
        /// <returns>The extracted template</returns>
        Task<ProvisioningTemplate> GetTemplateAsync(ExtractConfiguration configuration = null);

        /// <summary>
        /// Extracts a provisioning template from the site this manager was obtained from.
        /// </summary>
        /// <param name="configuration">What to extract and how, or <c>null</c> for everything</param>
        /// <returns>The extracted template</returns>
        ProvisioningTemplate GetTemplate(ExtractConfiguration configuration = null);

        /// <summary>
        /// Applies a tenant template - a hierarchy of sites and tenant level settings.
        /// </summary>
        /// <param name="hierarchy">The hierarchy to apply</param>
        /// <param name="sequenceId">Which sequence within the hierarchy to apply</param>
        /// <param name="configuration">Which handlers to run and how, or <c>null</c> for all of them</param>
        /// <remarks>Needs a context pointed at the tenant admin site.</remarks>
        Task ApplyTenantTemplateAsync(ProvisioningHierarchy hierarchy, string sequenceId, ApplyConfiguration configuration = null);

        /// <summary>
        /// Applies a tenant template - a hierarchy of sites and tenant level settings.
        /// </summary>
        /// <param name="hierarchy">The hierarchy to apply</param>
        /// <param name="sequenceId">Which sequence within the hierarchy to apply</param>
        /// <param name="configuration">Which handlers to run and how, or <c>null</c> for all of them</param>
        /// <remarks>Needs a context pointed at the tenant admin site.</remarks>
        void ApplyTenantTemplate(ProvisioningHierarchy hierarchy, string sequenceId, ApplyConfiguration configuration = null);

        /// <summary>
        /// Extracts a tenant template - a hierarchy of sites and tenant level settings.
        /// </summary>
        /// <param name="configuration">What to extract and how, or <c>null</c> for everything</param>
        /// <returns>The extracted hierarchy</returns>
        /// <remarks>Needs a context pointed at the tenant admin site.</remarks>
        Task<ProvisioningHierarchy> GetTenantTemplateAsync(ExtractConfiguration configuration = null);

        /// <summary>
        /// Extracts a tenant template - a hierarchy of sites and tenant level settings.
        /// </summary>
        /// <param name="configuration">What to extract and how, or <c>null</c> for everything</param>
        /// <returns>The extracted hierarchy</returns>
        /// <remarks>Needs a context pointed at the tenant admin site.</remarks>
        ProvisioningHierarchy GetTenantTemplate(ExtractConfiguration configuration = null);
    }
}
