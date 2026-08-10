using PnP.Core.Services.Core.CSOM;
using PnP.Core.Services.Core.CSOM.QueryAction;
using PnP.Core.Services.Core.CSOM.QueryIdentities;
using PnP.Core.Services.Core.CSOM.Requests;
using PnP.Core.Services.Core.CSOM.Utils;
using System;
using System.Collections.Generic;

namespace PnP.Core.Provisioning.Services.Core.CSOM.Requests.ContentTypes
{
    /// <summary>
    /// Sets a content type's properties and commits them, optionally pushing the change to child
    /// content types.
    /// </summary>
    /// <remarks>
    /// <para><b>Two reasons this exists rather than <c>IContentType.UpdateAsync()</c>.</b></para>
    /// <para>1. <b>PnP Core cannot update a content type read from a collection.</b> Updating the
    /// object <c>ContentTypes.AddAsync</c> returns works; updating one obtained from
    /// <c>Web.ContentTypes</c> fails with <i>"An unexpected 'PrimitiveValue' node was found when
    /// reading from the JSON reader"</i> - the request it builds is malformed. PnP Core's own tests
    /// only cover the first case, which is why the second is broken. A provisioning run is almost
    /// entirely the second case: the content type usually already exists.</para>
    /// <para>2. <b><c>UpdateChildren</c> cannot be expressed at all.</b> CSOM's
    /// <c>ContentType.Update(bool)</c> decides whether the change is pushed down to derived content
    /// types and to the list content types that inherit from it. The template controls that per
    /// content type; <c>UpdateAsync()</c> takes no argument, so the choice would be lost.</para>
    /// <para>Only properties explicitly added are sent, so this never overwrites a property the
    /// template did not mention.</para>
    /// </remarks>
    internal sealed class UpdateContentTypeRequest : IRequest<object>
    {
        private readonly string contentTypeIdentity;
        private readonly bool updateChildren;
        private readonly List<(string Name, string Type, object Value)> properties = new List<(string, string, object)>();

        /// <summary>
        /// Creates the request.
        /// </summary>
        /// <param name="siteId">The site collection id</param>
        /// <param name="webId">The web id</param>
        /// <param name="contentTypeId">The content type's string id</param>
        /// <param name="updateChildren">Whether to push the change to derived content types</param>
        internal UpdateContentTypeRequest(Guid siteId, Guid webId, string contentTypeId, bool updateChildren)
        {
            if (string.IsNullOrEmpty(contentTypeId))
            {
                throw new ArgumentException("A content type id is required.", nameof(contentTypeId));
            }

            contentTypeIdentity = CsomIdentity.ContentType(siteId, webId, contentTypeId);
            this.updateChildren = updateChildren;
        }

        /// <summary>Whether any property was queued.</summary>
        internal bool HasChanges => properties.Count > 0;

        internal void SetString(string name, string value)
        {
            properties.Add((name, "String", value));
        }

        internal void SetBoolean(string name, bool value)
        {
            properties.Add((name, "Boolean", value));
        }

        public object Result { get; private set; }

        public List<ActionObjectPath> GetRequest(IIdProvider idProvider)
        {
            var result = new List<ActionObjectPath>();

            int contentTypeId = idProvider.GetActionId();
            result.Add(new ActionObjectPath
            {
                ObjectPath = new Identity { Id = contentTypeId, Name = contentTypeIdentity },
            });

            foreach ((string name, string type, object value) in properties)
            {
                result.Add(new ActionObjectPath
                {
                    Action = new SetPropertyAction
                    {
                        Id = idProvider.GetActionId(),
                        ObjectPathId = contentTypeId.ToString(),
                        Name = name,
                        SetParameter = new Parameter { Type = type, Value = value },
                    },
                });
            }

            // Every setter above only stages. Without this the call succeeds and changes nothing -
            // the same trap as SetValueForUICulture.
            result.Add(new ActionObjectPath
            {
                Action = new MethodAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = contentTypeId.ToString(),
                    Name = "Update",
                    Parameters = new List<Parameter> { new Parameter { Type = "Boolean", Value = updateChildren } },
                },
            });

            return result;
        }

        public void ProcessResponse(string response)
        {
            // Nothing to read back.
        }
    }
}
