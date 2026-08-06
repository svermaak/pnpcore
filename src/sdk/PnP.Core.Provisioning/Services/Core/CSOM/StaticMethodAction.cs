using PnP.Core.Services.Core.CSOM.QueryAction;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Core.Provisioning.Services.Core.CSOM
{
    /// <summary>
    /// Invokes a void-returning static CSOM method as an action.
    /// </summary>
    /// <remarks>
    /// <para>PnP Core models a static method that <em>returns</em> something as a
    /// <c>StaticMethodPath</c> in <c>&lt;ObjectPaths&gt;</c> plus an <c>&lt;ObjectPath&gt;</c>
    /// action referencing it. That shape does not work for a <c>void</c> static: there is no
    /// resulting object to path to, and SharePoint expects the call itself in
    /// <c>&lt;Actions&gt;</c> as a <c>&lt;StaticMethod&gt;</c> element.</para>
    /// <para>PnP Core has no action type for that case - none of its 17 built-in requests needs
    /// one - so it is added here rather than in <c>PnP.Core</c>. Provisioning has several:
    /// <c>ProjectPolicy.ApplyProjectPolicy</c>, <c>SiteImageRenditions.SetRenditions</c>.</para>
    /// <para>Serialization matches <c>MethodAction</c>'s, swapping <c>ObjectPathId</c> for
    /// <c>TypeId</c>.</para>
    /// </remarks>
    internal sealed class StaticMethodAction : BaseAction
    {
        /// <summary>
        /// The method name.
        /// </summary>
        internal string Name { get; set; }

        /// <summary>
        /// The declaring type's CSOM server type id - see <see cref="CsomTypeIds"/>.
        /// </summary>
        internal string TypeId { get; set; }

        /// <summary>
        /// The method's parameters, excluding the implicit <c>ClientRuntimeContext</c>.
        /// </summary>
        internal List<Parameter> Parameters { get; set; }

        public override string ToString()
        {
            string parametersPart = string.Empty;

            if (Parameters != null && Parameters.Count > 0)
            {
                parametersPart = $"<Parameters>{string.Join("", Parameters.Select(p => p.SerializeParameter()))}</Parameters>";
            }

            return $"<StaticMethod Id=\"{Id}\" Name=\"{Name}\" TypeId=\"{TypeId}\">{parametersPart}</StaticMethod>";
        }
    }
}
