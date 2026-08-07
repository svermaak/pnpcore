using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.ObjectHandlers.Utilities
{
    /// <summary>
    /// Loads a modern page by the name a template uses for it.
    /// </summary>
    /// <remarks>
    /// <para>A template names a page relative to the pages library, folder included -
    /// <c>Archive/Home.aspx</c>. <see cref="IWeb.GetPagesAsync"/> matches on the <em>file name</em>,
    /// so that call returns both <c>Home.aspx</c> and <c>Archive/Home.aspx</c>. Taking the first hit
    /// would silently load the wrong page - and then overwrite it.</para>
    /// <para>Shared by the apply and extract handlers so they cannot disagree about which page a
    /// name refers to.</para>
    /// </remarks>
    internal static class PageLookup
    {
        /// <summary>
        /// Finds the page a library-relative name refers to.
        /// </summary>
        /// <returns>The page, or <c>null</c> when the site has no such page</returns>
        internal static async Task<IPage> LoadAsync(PnPContext context, string pageName)
        {
            if (string.IsNullOrEmpty(pageName))
            {
                return null;
            }

            string fileName = System.IO.Path.GetFileName(pageName);
            List<IPage> pages = await context.Web.GetPagesAsync(fileName).ConfigureAwait(false);

            if (pages == null || pages.Count == 0)
            {
                return null;
            }

            if (pages.Count == 1)
            {
                return pages[0];
            }

            string folder = System.IO.Path.GetDirectoryName(pageName)?.Replace("\\", "/").Trim('/') ?? string.Empty;

            return pages.FirstOrDefault(p =>
                       string.Equals(p.Folder?.Trim('/') ?? string.Empty, folder, StringComparison.OrdinalIgnoreCase))
                   ?? pages[0];
        }
    }
}
