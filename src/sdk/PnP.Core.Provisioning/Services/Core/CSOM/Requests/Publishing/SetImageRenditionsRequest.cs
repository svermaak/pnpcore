using PnP.Core.Provisioning.Model;
using PnP.Core.Services.Core.CSOM;
using PnP.Core.Services.Core.CSOM.QueryAction;
using PnP.Core.Services.Core.CSOM.QueryIdentities;
using PnP.Core.Services.Core.CSOM.Requests;
using PnP.Core.Services.Core.CSOM.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Core.Provisioning.Services.Core.CSOM.Requests.Publishing
{
    /// <summary>
    /// Replaces the site collection's image renditions.
    /// </summary>
    /// <remarks>
    /// <para>Backlog <b>T20</b>. <c>SiteImageRenditions.SetRenditions</c> takes the complete list,
    /// not a delta - so the caller must read the current renditions, merge, and send the whole set
    /// back. Sending only the new ones deletes the rest.</para>
    /// <para>Each rendition is built by a <c>Constructor</c> object path and populated with
    /// <c>SetProperty</c> actions, because <c>ImageRendition</c> is a client object rather than a
    /// value type - there is no way to express it as a single serialized parameter.</para>
    /// </remarks>
    internal sealed class SetImageRenditionsRequest : IRequest<object>
    {
        private readonly IReadOnlyList<ImageRenditionInfo> renditions;

        internal SetImageRenditionsRequest(IReadOnlyList<ImageRenditionInfo> renditions)
        {
            this.renditions = renditions ?? throw new ArgumentNullException(nameof(renditions));
        }

        public object Result { get; private set; }

        public List<ActionObjectPath> GetRequest(IIdProvider idProvider)
        {
            var result = new List<ActionObjectPath>();
            var renditionPathIds = new List<int>();

            foreach (ImageRenditionInfo rendition in renditions)
            {
                var constructor = new ConstructorPath
                {
                    Id = idProvider.GetActionId(),
                    TypeId = CsomTypeIds.ImageRendition,
                    Parameters = new MethodParameter { Properties = new List<Parameter>() }
                };

                result.Add(new ActionObjectPath
                {
                    Action = new BaseAction
                    {
                        Id = idProvider.GetActionId(),
                        ObjectPathId = constructor.Id.ToString()
                    },
                    ObjectPath = constructor
                });

                // Only Name, Width and Height are settable; Id and Version are server-assigned.
                result.Add(SetProperty(idProvider, constructor.Id, "Name", "String", rendition.Name));
                result.Add(SetProperty(idProvider, constructor.Id, "Width", "Number", rendition.Width));
                result.Add(SetProperty(idProvider, constructor.Id, "Height", "Number", rendition.Height));

                renditionPathIds.Add(constructor.Id);
            }

            result.Add(new ActionObjectPath
            {
                Action = new StaticMethodAction
                {
                    Id = idProvider.GetActionId(),
                    TypeId = CsomTypeIds.SiteImageRenditions,
                    Name = "SetRenditions",
                    Parameters = new List<Parameter>
                    {
                        new ObjectArrayParameter
                        {
                            ObjectPathIds = renditionPathIds
                        }
                    }
                }
            });

            return result;
        }

        private static ActionObjectPath SetProperty(IIdProvider idProvider, int objectPathId, string name, string type, object value)
        {
            return new ActionObjectPath
            {
                Action = new SetPropertyAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = objectPathId.ToString(),
                    Name = name,
                    SetParameter = new Parameter { Type = type, Value = value }
                }
            };
        }

        public void ProcessResponse(string response)
        {
            // Nothing to read back.
        }
    }
}
