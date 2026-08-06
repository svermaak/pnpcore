namespace PnP.Core.Provisioning.Model
{
    /// <summary>
    /// A site policy as reported by the information management policy CSOM API.
    /// </summary>
    /// <remarks>
    /// A plain DTO rather than a PnP Core model: site policies are read-only from the provisioning
    /// engine's point of view - a template names one and the engine applies it - so there is
    /// nothing for a full data model to buy.
    /// </remarks>
    internal sealed class SitePolicyInfo
    {
        /// <summary>
        /// The policy's name, which is what a template's <c>&lt;pnp:SitePolicy&gt;</c> element carries.
        /// </summary>
        internal string Name { get; set; }

        /// <summary>
        /// The policy's description.
        /// </summary>
        internal string Description { get; set; }
    }
}
