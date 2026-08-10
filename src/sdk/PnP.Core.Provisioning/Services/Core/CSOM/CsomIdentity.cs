using System;
using System.Globalization;

namespace PnP.Core.Provisioning.Services.Core.CSOM
{
    /// <summary>
    /// Builds the CSOM object identity strings that address an already-existing server object.
    /// </summary>
    /// <remarks>
    /// <para>A CSOM request reaches an existing object one of two ways: by walking an object path
    /// from a static method, or by naming its <em>identity</em> directly. The identity is a colon
    /// delimited string, and the shape is stable enough to construct:</para>
    /// <code>
    /// &lt;session&gt;|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:&lt;siteId&gt;:web:&lt;webId&gt;[:list:&lt;listId&gt;][:...]
    /// </code>
    /// <para>The leading segment is a per-session correlation prefix that the server echoes back
    /// but does not validate on input - PnP Core's own requests hardcode arbitrary values for it,
    /// and so does this. The <c>740c6a0b-…</c> GUID is the SharePoint client object model's
    /// application id and is fixed.</para>
    /// </remarks>
    internal static class CsomIdentity
    {
        /// <summary>
        /// The SharePoint client object model application id. Fixed; part of every identity.
        /// </summary>
        private const string ApplicationId = "740c6a0b-85e2-48a0-a494-e0f1759d4aa7";

        /// <summary>
        /// A correlation prefix. The server does not validate it on input.
        /// </summary>
        private const string SessionPrefix = "00000000-0000-0000-0000-000000000000";

        /// <summary>
        /// The identity of a web.
        /// </summary>
        internal static string Web(Guid siteId, Guid webId)
        {
            return string.Format(CultureInfo.InvariantCulture,
                "{0}|{1}:site:{2}:web:{3}", SessionPrefix, ApplicationId, siteId, webId);
        }

        /// <summary>
        /// The identity of a site collection.
        /// </summary>
        /// <remarks>
        /// A site identity still carries the <c>:web:</c> segment - SharePoint addresses the site
        /// through the web that was used to reach it.
        /// </remarks>
        internal static string Site(Guid siteId, Guid webId)
        {
            return string.Format(CultureInfo.InvariantCulture,
                "{0}|{1}:site:{2}:web:{3}:site:{2}", SessionPrefix, ApplicationId, siteId, webId);
        }

        /// <summary>
        /// The identity of a list.
        /// </summary>
        internal static string List(Guid siteId, Guid webId, Guid listId)
        {
            return string.Format(CultureInfo.InvariantCulture,
                "{0}|{1}:site:{2}:web:{3}:list:{4}", SessionPrefix, ApplicationId, siteId, webId, listId);
        }

        /// <summary>
        /// The identity of a field, at web scope.
        /// </summary>
        internal static string Field(Guid siteId, Guid webId, Guid fieldId)
        {
            return string.Format(CultureInfo.InvariantCulture,
                "{0}|{1}:site:{2}:web:{3}:field:{4}", SessionPrefix, ApplicationId, siteId, webId, fieldId);
        }

        /// <summary>
        /// The identity of a field, at list scope.
        /// </summary>
        internal static string ListField(Guid siteId, Guid webId, Guid listId, Guid fieldId)
        {
            return string.Format(CultureInfo.InvariantCulture,
                "{0}|{1}:site:{2}:web:{3}:list:{4}:field:{5}", SessionPrefix, ApplicationId, siteId, webId, listId, fieldId);
        }

        /// <summary>
        /// The identity of a content type, at web scope.
        /// </summary>
        internal static string ContentType(Guid siteId, Guid webId, string contentTypeId)
        {
            return string.Format(CultureInfo.InvariantCulture,
                "{0}|{1}:site:{2}:web:{3}:contenttype:{4}", SessionPrefix, ApplicationId, siteId, webId, contentTypeId);
        }

        /// <summary>
        /// The identity of a user custom action, at web scope.
        /// </summary>
        internal static string UserCustomAction(Guid siteId, Guid webId, Guid customActionId)
        {
            return string.Format(CultureInfo.InvariantCulture,
                "{0}|{1}:site:{2}:web:{3}:useraction:{4}", SessionPrefix, ApplicationId, siteId, webId, customActionId);
        }

        /// <summary>
        /// The identity of a user custom action, at site collection scope.
        /// </summary>
        /// <remarks>
        /// Site and web scoped actions have <em>different</em> identities even when the ids look
        /// alike, and addressing a site-scoped action through the web path resolves to nothing -
        /// which SharePoint reports as a null object rather than an error.
        /// </remarks>
        internal static string SiteUserCustomAction(Guid siteId, Guid webId, Guid customActionId)
        {
            return string.Format(CultureInfo.InvariantCulture,
                "{0}:useraction:{1}", Site(siteId, webId), customActionId);
        }
    }
}
