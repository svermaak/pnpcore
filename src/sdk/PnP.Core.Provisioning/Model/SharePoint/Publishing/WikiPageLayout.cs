namespace PnP.Core.Provisioning.Model
{
    /// <summary>
    /// Out of the box wiki page layouts enumeration
    /// </summary>
    /// <remarks>
    /// Ported from <c>PnP.Framework.WikiPageLayout</c>. Despite describing a classic SharePoint concept
    /// this was never a CSOM type - PnP Framework declared it itself - so it moves across unchanged.
    /// It is consumed by the classic wiki page model (<see cref="Page"/>) and by the generated
    /// provisioning schema classes, which declare a matching enum per schema version.
    /// </remarks>
    public enum WikiPageLayout
    {
        /// <summary>
        /// One column layout
        /// </summary>
        OneColumn = 0,

        /// <summary>
        /// One column with sidebar layout
        /// </summary>
        OneColumnSideBar = 1,

        /// <summary>
        /// Two columns layout
        /// </summary>
        TwoColumns = 2,

        /// <summary>
        /// Two columns with header layout
        /// </summary>
        TwoColumnsHeader = 3,

        /// <summary>
        /// Two columns with header and footer layout
        /// </summary>
        TwoColumnsHeaderFooter = 4,

        /// <summary>
        /// Three columns layout
        /// </summary>
        ThreeColumns = 5,

        /// <summary>
        /// Three columns with header layout
        /// </summary>
        ThreeColumnsHeader = 6,

        /// <summary>
        /// Three columns with header and footer layout
        /// </summary>
        ThreeColumnsHeaderFooter = 7,

        /// <summary>
        /// Custom layout
        /// </summary>
        Custom = 8
    }
}
