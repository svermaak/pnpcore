using PnP.Core.Provisioning.Model;

namespace PnP.Core.Provisioning.Services.Core.CSOM.Requests.Audit
{
    /// <summary>
    /// The site collection's audit configuration, as read back from CSOM.
    /// </summary>
    /// <remarks>
    /// <para>A class rather than a bare <c>AuditMaskType?</c> because a CSOM request's result type
    /// must be a <b>reference type</b>. <c>IRequest&lt;out T&gt;</c> is covariant, and PnP Core's
    /// pipeline casts every request to <c>IRequest&lt;object&gt;</c> to batch them - a cast that
    /// only succeeds for reference types. A value-typed result compiles fine and then throws
    /// <c>InvalidCastException</c> at run time, inside the SDK, with nothing pointing back at the
    /// request that caused it.</para>
    /// <para>None of PnP Core's own 17 requests returns a value type, so the constraint is
    /// undocumented there. <c>CsomRequestPayloadTests</c> asserts it across the whole assembly.</para>
    /// </remarks>
    internal sealed class AuditSettingsInfo
    {
        /// <summary>
        /// The events being audited.
        /// </summary>
        internal AuditMaskType AuditFlags { get; set; }
    }
}
