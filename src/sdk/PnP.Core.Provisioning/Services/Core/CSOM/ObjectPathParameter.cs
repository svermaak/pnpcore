using PnP.Core.Services.Core.CSOM.QueryAction;
using System.Globalization;

namespace PnP.Core.Provisioning.Services.Core.CSOM
{
    /// <summary>
    /// A parameter that refers to another object path in the same request, rather than carrying a
    /// literal value.
    /// </summary>
    /// <remarks>
    /// <para>Decision <b>D12</b>: a CSOM building block PnP Core does not have, added here rather
    /// than upstream to keep the change local.</para>
    /// <para>It is needed whenever a property is set to a <em>server object</em> instead of a
    /// scalar - <c>Web.AssociatedOwnerGroup = &lt;some group&gt;</c> being the case that forced it.
    /// PnP Core's <see cref="Parameter"/> serializes a value with a type name; the wire form for an
    /// object reference has no value and no type at all, only the id of the path that produced the
    /// object:</para>
    /// <code>&lt;Parameter ObjectPathId="14" /&gt;</code>
    /// <para>None of PnP Core's own 17 requests assigns an object to a property, which is why the
    /// shape is missing.</para>
    /// </remarks>
    internal sealed class ObjectPathParameter : Parameter
    {
        /// <summary>
        /// The id of the object path this parameter refers to.
        /// </summary>
        internal int ReferencedObjectPathId { get; set; }

        internal override string SerializeParameter()
        {
            return string.Format(CultureInfo.InvariantCulture,
                "<{0} ObjectPathId=\"{1}\" />", ParameterTagName, ReferencedObjectPathId);
        }
    }
}
