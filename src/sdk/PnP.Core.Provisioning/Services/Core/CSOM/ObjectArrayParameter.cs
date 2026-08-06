using PnP.Core.Services.Core.CSOM.QueryAction;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Core.Provisioning.Services.Core.CSOM
{
    /// <summary>
    /// An array parameter whose elements are references to other object paths in the same request.
    /// </summary>
    /// <remarks>
    /// <para>PnP Core's <see cref="Parameter"/> serializes arrays of <em>values</em> -
    /// <c>List&lt;string&gt;</c>, <c>List&lt;Guid&gt;</c>, <c>List&lt;int&gt;</c> - and
    /// <c>ObjectReferenceParameter</c> serializes a single object reference. Neither covers an
    /// array <em>of</em> object references, which is what
    /// <c>SiteImageRenditions.SetRenditions(IList&lt;ImageRendition&gt;)</c> needs: each rendition
    /// is a client object built by its own constructor path.</para>
    /// <para>Added here rather than in <c>PnP.Core</c> since no built-in request needs it.</para>
    /// </remarks>
    internal sealed class ObjectArrayParameter : Parameter
    {
        /// <summary>
        /// The object path ids of the array's elements, in order.
        /// </summary>
        internal List<int> ObjectPathIds { get; set; } = new List<int>();

        internal override string SerializeParameter()
        {
            string elements = string.Join("", ObjectPathIds.Select(id => $"<Object ObjectPathId=\"{id}\" />"));
            return $"<Parameter Type=\"Array\">{elements}</Parameter>";
        }
    }
}
