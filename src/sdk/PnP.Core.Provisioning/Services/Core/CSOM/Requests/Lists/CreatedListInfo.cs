using System;

namespace PnP.Core.Provisioning.Services.Core.CSOM.Requests.Lists
{
    /// <summary>
    /// What <see cref="CreateListRequest"/> reads back about the list it created.
    /// </summary>
    /// <remarks>
    /// A class wrapping a single <see cref="Guid"/> looks like ceremony, and is not: PnP Core
    /// batches requests by casting them to <c>IRequest&lt;object&gt;</c>, <c>IRequest&lt;out T&gt;</c>
    /// is covariant, and covariance does not apply to value types. <c>IRequest&lt;Guid&gt;</c>
    /// compiles cleanly and then throws <c>InvalidCastException</c> from inside the SDK.
    /// <c>EveryCsomRequestResultTypeIsAReferenceType</c> guards this for the whole assembly.
    /// </remarks>
    internal sealed class CreatedListInfo
    {
        /// <summary>The id of the newly created list.</summary>
        internal Guid Id { get; set; }
    }
}
