using PnP.Core.Services.Core.CSOM;
using PnP.Core.Services.Core.CSOM.QueryAction;
using PnP.Core.Services.Core.CSOM.QueryIdentities;
using PnP.Core.Services.Core.CSOM.Requests;
using PnP.Core.Services.Core.CSOM.Utils;
using System;
using System.Collections.Generic;

namespace PnP.Core.Provisioning.Services.Core.CSOM.Requests.UserResources
{
    /// <summary>
    /// Sets a localized value for one culture on one resource property.
    /// </summary>
    /// <remarks>
    /// <para>Backlog <b>T6</b> - the cross-cutting one. PnP Core has <c>IUserResource</c> support
    /// <em>commented out</em> across <c>IContentType</c>, <c>IField</c>, <c>IUserCustomAction</c>
    /// and <c>IWeb</c>: a known deferral, not an oversight. Until it lands there, provisioning
    /// reaches <c>SP.UserResource.SetValueForUICulture</c> directly.</para>
    /// <para>A dozen handlers need this - every <c>Title</c> and <c>Description</c> that a template
    /// can localize - which is why <b>D6</b> moved T6 to the front of phase 6 and made this request
    /// a required phase 4 deliverable rather than an optional one.</para>
    /// <para>Several cultures on one object should go in a single batch via
    /// <c>CsomRequestSender.SendManyAsync</c>: each culture is its own <c>SetValueForUICulture</c>
    /// call, and a template that localizes into eight languages should not cost eight round trips.</para>
    /// </remarks>
    internal sealed class SetValueForUICultureRequest : IRequest<object>
    {
        private readonly UserResourcePath resource;
        private readonly string cultureName;
        private readonly string value;

        /// <summary>
        /// Creates the request.
        /// </summary>
        /// <param name="resource">Which resource property on which object</param>
        /// <param name="cultureName">A culture name such as <c>en-US</c> - not an LCID</param>
        /// <param name="value">The localized text</param>
        internal SetValueForUICultureRequest(UserResourcePath resource, string cultureName, string value)
        {
            this.resource = resource ?? throw new ArgumentNullException(nameof(resource));

            if (string.IsNullOrEmpty(cultureName))
            {
                throw new ArgumentException("A culture name is required.", nameof(cultureName));
            }

            this.cultureName = cultureName;
            this.value = value;
        }

        public object Result { get; private set; }

        public List<ActionObjectPath> GetRequest(IIdProvider idProvider)
        {
            var result = new List<ActionObjectPath>();

            int parentIdentityId = resource.AppendParentPath(idProvider, result);

            // e.g. Web.TitleResource
            int resourcePropertyId = idProvider.GetActionId();
            result.Add(new ActionObjectPath
            {
                ObjectPath = new Property
                {
                    Id = resourcePropertyId,
                    ParentId = parentIdentityId,
                    Name = resource.PropertyName
                }
            });

            result.Add(new ActionObjectPath
            {
                Action = new MethodAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = resourcePropertyId.ToString(),
                    Name = "SetValueForUICulture",
                    Parameters = new List<Parameter>
                    {
                        new Parameter { Type = "String", Value = cultureName },
                        new Parameter { Type = "String", Value = value }
                    }
                }
            });

            // Persist. SetValueForUICulture only STAGES the change - without this the call succeeds,
            // raises nothing, and the value is silently discarded. The first live run of this
            // request failed exactly that way against a site with 51 languages enabled, and the
            // payload tests could not have caught it because the XML was structurally correct.
            //
            // The method and its argument differ by parent type; UserResourcePath carries which.
            result.Add(new ActionObjectPath
            {
                Action = new MethodAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = parentIdentityId.ToString(),
                    Name = resource.ParentUpdateMethod,
                    Parameters = resource.ParentUpdateTakesFlag
                        ? new List<Parameter> { new Parameter { Type = "Boolean", Value = true } }
                        : new List<Parameter>()
                }
            });

            return result;
        }

        public void ProcessResponse(string response)
        {
            // Nothing to read back.
        }
    }
}
