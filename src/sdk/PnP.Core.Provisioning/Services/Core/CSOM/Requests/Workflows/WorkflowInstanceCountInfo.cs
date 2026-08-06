namespace PnP.Core.Provisioning.Model
{
    /// <summary>
    /// The number of running instances of a workflow association.
    /// </summary>
    /// <remarks>
    /// A class rather than a bare <c>int</c> because a CSOM request's result type must be a
    /// reference type - see <c>AuditSettingsInfo</c>.
    /// </remarks>
    internal sealed class WorkflowInstanceCountInfo
    {
        /// <summary>
        /// How many instances are currently running.
        /// </summary>
        internal int Count { get; set; }
    }
}