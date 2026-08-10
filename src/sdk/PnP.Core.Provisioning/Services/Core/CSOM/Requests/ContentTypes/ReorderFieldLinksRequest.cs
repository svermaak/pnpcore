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
    /// Puts a content type's columns into the order a template specifies.
    /// </summary>
    /// <remarks>
    /// <para><b>Backlog T15-adjacent, and a real gap.</b> PnP Core models field links well enough to
    /// add and configure them but has no <c>Reorder</c> - <c>IList.ReorderContentTypesAsync</c>
    /// exists, the equivalent for a content type's columns does not. CSOM's
    /// <c>ContentType.FieldLinks.Reorder(string[])</c> is the only way.</para>
    /// <para>It is not cosmetic. Field order is the order columns appear on the new, edit and display
    /// forms, and a content type that inherits from another gets the parent's columns first with the
    /// template's own appended - so without this, every derived content type comes out in the wrong
    /// order regardless of what the template said.</para>
    /// <para>Reorder only <em>arranges</em>; it does not add or remove. The names passed must already
    /// be links on the content type, so this runs after the field links have been created.</para>
    /// </remarks>
    internal sealed class ReorderFieldLinksRequest : IRequest<object>
    {
        private readonly string contentTypeIdentity;
        private readonly List<string> fieldInternalNames;

        /// <summary>
        /// Creates the request.
        /// </summary>
        /// <param name="siteId">The site collection id</param>
        /// <param name="webId">The web id</param>
        /// <param name="contentTypeId">The content type's string id, for example <c>0x0100...</c></param>
        /// <param name="fieldInternalNames">The internal names, in the wanted order</param>
        internal ReorderFieldLinksRequest(Guid siteId, Guid webId, string contentTypeId, List<string> fieldInternalNames)
        {
            if (string.IsNullOrEmpty(contentTypeId))
            {
                throw new ArgumentException("A content type id is required.", nameof(contentTypeId));
            }

            if (fieldInternalNames == null || fieldInternalNames.Count == 0)
            {
                throw new ArgumentException("At least one field name is required.", nameof(fieldInternalNames));
            }

            contentTypeIdentity = CsomIdentity.ContentType(siteId, webId, contentTypeId);
            this.fieldInternalNames = fieldInternalNames;
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

            int fieldLinksId = idProvider.GetActionId();
            result.Add(new ActionObjectPath
            {
                ObjectPath = new Property { Id = fieldLinksId, ParentId = contentTypeId, Name = "FieldLinks" },
            });

            result.Add(new ActionObjectPath
            {
                Action = new MethodAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = fieldLinksId.ToString(),
                    Name = "Reorder",

                    // A List<string> serialises as the <Parameter Type="Array"> CSOM expects.
                    Parameters = new List<Parameter> { new Parameter { Type = "String", Value = fieldInternalNames } },
                },
            });

            // Reorder stages the change like every other CSOM setter; without the update it is
            // accepted and discarded. False, because column order is not pushed to child content
            // types - PnP Framework passed the content type's own UpdateChildren for the property
            // update and never for this one.
            result.Add(new ActionObjectPath
            {
                Action = new MethodAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = contentTypeId.ToString(),
                    Name = "Update",
                    Parameters = new List<Parameter> { new Parameter { Type = "Boolean", Value = false } },
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
