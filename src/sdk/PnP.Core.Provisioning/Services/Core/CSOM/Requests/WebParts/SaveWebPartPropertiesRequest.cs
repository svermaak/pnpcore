using PnP.Core.Provisioning.Model;
using PnP.Core.Services.Core.CSOM;
using PnP.Core.Services.Core.CSOM.QueryAction;
using PnP.Core.Services.Core.CSOM.QueryIdentities;
using PnP.Core.Services.Core.CSOM.Requests;
using PnP.Core.Services.Core.CSOM.Utils;
using System;
using System.Collections.Generic;
using System.Text.Json;

namespace PnP.Core.Provisioning.Services.Core.CSOM.Requests.WebParts
{    /// <summary>
    /// Updates a web part's properties.
    /// </summary>
    /// <remarks>
    /// <para>Backlog <b>T18</b>. Sets properties on <c>WebPartDefinition.WebPart</c> and then calls
    /// <c>SaveWebPartChanges()</c> - the setters alone stage the change without persisting it, the
    /// same trap as <c>Audit.Update()</c> and the term store's <c>CommitAll()</c>.</para>
    /// <para>Only the strongly typed properties are settable this way. Arbitrary web part
    /// properties live in <c>WebPart.Properties</c>, which is a property bag and needs
    /// <c>SetFieldValue</c> calls instead - that path arrives with <c>ObjectPages</c> in phase 7,
    /// which is the only consumer that needs it.</para>
    /// <para><b><c>ZoneIndex</c> is deliberately absent.</b> A first version set it here and
    /// SharePoint rejected the whole request with <i>Field or property "ZoneIndex" does not
    /// exist</i>. The CSOM assembly confirms why: on <c>SP.WebParts.WebPart</c>, <c>Title</c>,
    /// <c>TitleUrl</c>, <c>Hidden</c> and <c>ExportMode</c> carry setters, but <c>ZoneIndex</c> is
    /// <b>read-only</b>. A web part is repositioned through
    /// <c>WebPartDefinition.MoveWebPartTo(zoneId, zoneIndex)</c> - see
    /// <see cref="MoveWebPartToRequest"/>. Position and properties are two separate operations in
    /// CSOM, and conflating them fails the request as a whole rather than just that one property.</para>
    /// </remarks>
    internal sealed class SaveWebPartPropertiesRequest : WebPartRequestBase, IRequest<object>
    {
        private readonly Guid siteId;
        private readonly Guid webId;
        private readonly string serverRelativeFileUrl;
        private readonly Guid webPartId;
        private readonly string title;

        internal SaveWebPartPropertiesRequest(Guid siteId, Guid webId, string serverRelativeFileUrl, Guid webPartId,
            string title = null)
        {
            if (string.IsNullOrEmpty(serverRelativeFileUrl))
            {
                throw new ArgumentException("A server relative file url is required.", nameof(serverRelativeFileUrl));
            }

            this.siteId = siteId;
            this.webId = webId;
            this.serverRelativeFileUrl = serverRelativeFileUrl;
            this.webPartId = webPartId;
            this.title = title;
        }

        public object Result { get; private set; }

        public List<ActionObjectPath> GetRequest(IIdProvider idProvider)
        {
            var paths = new List<ActionObjectPath>();

            int managerId = EmitWebPartManager(idProvider, paths, siteId, webId, serverRelativeFileUrl);
            int webPartsId = idProvider.GetActionId();

            paths.Add(new ActionObjectPath
            {
                ObjectPath = new Property
                {
                    Id = webPartsId,
                    ParentId = managerId,
                    Name = "WebParts"
                }
            });

            int definitionId = EmitMethod(idProvider, paths, webPartsId, "GetById",
                new Parameter { Type = "Guid", Value = webPartId });

            int webPartId2 = idProvider.GetActionId();
            paths.Add(new ActionObjectPath
            {
                ObjectPath = new Property
                {
                    Id = webPartId2,
                    ParentId = definitionId,
                    Name = "WebPart"
                }
            });

            if (title != null)
            {
                paths.Add(new ActionObjectPath
                {
                    Action = new SetPropertyAction
                    {
                        Id = idProvider.GetActionId(),
                        ObjectPathId = webPartId2.ToString(),
                        Name = "Title",
                        SetParameter = new Parameter { Type = "String", Value = title }
                    }
                });
            }

            paths.Add(new ActionObjectPath
            {
                Action = new MethodAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = definitionId.ToString(),
                    Name = "SaveWebPartChanges",
                    Parameters = new List<Parameter>()
                }
            });

            return paths;
        }

        public void ProcessResponse(string response)
        {
        }
    }
}