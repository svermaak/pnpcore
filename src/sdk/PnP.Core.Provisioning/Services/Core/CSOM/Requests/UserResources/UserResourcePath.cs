using System;

namespace PnP.Core.Provisioning.Services.Core.CSOM.Requests.UserResources
{
    /// <summary>
    /// Identifies one localizable text property on one SharePoint object, and knows how to persist
    /// a change to it.
    /// </summary>
    /// <remarks>
    /// <para>Backlog <b>T6</b>. A <c>UserResource</c> is not addressable on its own - it hangs off a
    /// parent as a named property (<c>Web.TitleResource</c>, <c>Field.DescriptionResource</c>,
    /// <c>ContentType.NameResource</c>, ...). So a request needs the parent's CSOM identity and the
    /// resource property's name.</para>
    /// <para><b>And a third thing, learned the hard way:</b> <c>SetValueForUICulture</c> only
    /// <em>stages</em> the change. Without an <c>Update</c> on the <b>parent</b> it is silently
    /// discarded - the call succeeds, no error is raised, and a read-back returns the object's
    /// default title as though nothing happened. The first live run of these requests failed
    /// exactly that way, on a site with 51 languages enabled, and no payload test could have caught
    /// it because the XML was structurally perfect.</para>
    /// <para>Worse, <b>the update method differs by parent type</b>, which is why this is encoded
    /// here rather than left to each call site:</para>
    /// <list type="table">
    /// <item><term>Web, List</term><description><c>Update()</c></description></item>
    /// <item><term>ContentType</term><description><c>Update(true)</c> - the flag pushes to children</description></item>
    /// <item><term>Field</term><description><c>UpdateAndPushChanges(true)</c></description></item>
    /// </list>
    /// <para>Use the factory methods rather than the constructor; they get this right.</para>
    /// </remarks>
    internal sealed class UserResourcePath
    {
        /// <summary>
        /// The CSOM identity of the object that owns the resource - see <see cref="CsomIdentity"/>.
        /// </summary>
        internal string ParentIdentity { get; }

        /// <summary>
        /// The name of the resource property on that object.
        /// </summary>
        internal string PropertyName { get; }

        /// <summary>
        /// The method that persists a staged change on the parent - <c>Update</c> or
        /// <c>UpdateAndPushChanges</c>.
        /// </summary>
        internal string ParentUpdateMethod { get; }

        /// <summary>
        /// Whether <see cref="ParentUpdateMethod"/> takes a boolean argument.
        /// </summary>
        internal bool ParentUpdateTakesFlag { get; }

        internal UserResourcePath(string parentIdentity, string propertyName,
            string parentUpdateMethod = "Update", bool parentUpdateTakesFlag = false)
        {
            if (string.IsNullOrEmpty(parentIdentity))
            {
                throw new ArgumentException("A parent identity is required.", nameof(parentIdentity));
            }

            if (string.IsNullOrEmpty(propertyName))
            {
                throw new ArgumentException("A resource property name is required.", nameof(propertyName));
            }

            ParentIdentity = parentIdentity;
            PropertyName = propertyName;
            ParentUpdateMethod = parentUpdateMethod;
            ParentUpdateTakesFlag = parentUpdateTakesFlag;
        }

        /// <summary>A localizable property on a web. Persisted with <c>Web.Update()</c>.</summary>
        internal static UserResourcePath ForWeb(Guid siteId, Guid webId, string propertyName)
        {
            return new UserResourcePath(CsomIdentity.Web(siteId, webId), propertyName);
        }

        /// <summary>A localizable property on a list. Persisted with <c>List.Update()</c>.</summary>
        internal static UserResourcePath ForList(Guid siteId, Guid webId, Guid listId, string propertyName)
        {
            return new UserResourcePath(CsomIdentity.List(siteId, webId, listId), propertyName);
        }

        /// <summary>
        /// A localizable property on a site column. Persisted with
        /// <c>Field.UpdateAndPushChanges(true)</c> so list columns derived from it follow.
        /// </summary>
        internal static UserResourcePath ForField(Guid siteId, Guid webId, Guid fieldId, string propertyName)
        {
            return new UserResourcePath(CsomIdentity.Field(siteId, webId, fieldId), propertyName,
                "UpdateAndPushChanges", parentUpdateTakesFlag: true);
        }

        /// <summary>
        /// A localizable property on a content type. Persisted with <c>ContentType.Update(true)</c>
        /// so child content types follow.
        /// </summary>
        internal static UserResourcePath ForContentType(Guid siteId, Guid webId, string contentTypeId, string propertyName)
        {
            return new UserResourcePath(CsomIdentity.ContentType(siteId, webId, contentTypeId), propertyName,
                "Update", parentUpdateTakesFlag: true);
        }

        /// <summary>
        /// A localizable property on a user custom action. Persisted with
        /// <c>UserCustomAction.Update()</c>.
        /// </summary>
        internal static UserResourcePath ForUserCustomAction(Guid siteId, Guid webId, Guid customActionId, string propertyName)
        {
            return new UserResourcePath(CsomIdentity.UserCustomAction(siteId, webId, customActionId), propertyName);
        }
    }

    /// <summary>
    /// The names of the localizable resource properties, per object type.
    /// </summary>
    /// <remarks>
    /// Mostly <c>TitleResource</c> / <c>DescriptionResource</c>, but a content type's display name
    /// property is <c>NameResource</c> - because the underlying property is <c>Name</c>, not
    /// <c>Title</c>. Getting that wrong silently fails to localize content types.
    /// </remarks>
    internal static class ResourceProperty
    {
        /// <summary>The title of a web, list, field, user custom action or navigation node.</summary>
        internal const string Title = "TitleResource";

        /// <summary>The description of a web, list, field, content type or user custom action.</summary>
        internal const string Description = "DescriptionResource";

        /// <summary>A content type's display name. Not <c>TitleResource</c>.</summary>
        internal const string ContentTypeName = "NameResource";
    }
}
