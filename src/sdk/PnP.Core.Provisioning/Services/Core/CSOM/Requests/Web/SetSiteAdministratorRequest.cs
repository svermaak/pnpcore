using PnP.Core.Services.Core.CSOM;
using PnP.Core.Services.Core.CSOM.QueryAction;
using PnP.Core.Services.Core.CSOM.QueryIdentities;
using PnP.Core.Services.Core.CSOM.Requests;
using PnP.Core.Services.Core.CSOM.Utils;
using System;
using System.Collections.Generic;

namespace PnP.Core.Provisioning.Services.Core.CSOM.Requests.Web
{
    /// <summary>
    /// Grants or revokes site collection administrator rights for one user.
    /// </summary>
    /// <remarks>
    /// <para><c>ISharePointUser.IsSiteAdmin</c> is <b>get-only</b> in PnP Core, and the REST entity
    /// does not accept a write to it either. A template's
    /// <c>&lt;pnp:AdditionalAdministrators&gt;</c> therefore needs CSOM.</para>
    /// <para>The user is reached by <c>Web</c> → <c>SiteUsers</c> → <c>GetById(id)</c> rather than
    /// by an identity string: a user identity is not one of the forms proven to work, and the
    /// object-path walk is what the CSOM client itself emits.</para>
    /// <para>The commit is <c>User.Update()</c>, on the user - not on the web.</para>
    /// </remarks>
    internal sealed class SetSiteAdministratorRequest : IRequest<object>
    {
        private readonly string webIdentity;
        private readonly int userId;
        private readonly bool isSiteAdmin;

        /// <summary>
        /// Creates the request.
        /// </summary>
        /// <param name="siteId">The site collection id</param>
        /// <param name="webId">The web id</param>
        /// <param name="userId">The site user id</param>
        /// <param name="isSiteAdmin">Whether the user should be a site collection administrator</param>
        internal SetSiteAdministratorRequest(Guid siteId, Guid webId, int userId, bool isSiteAdmin)
        {
            webIdentity = CsomIdentity.Web(siteId, webId);
            this.userId = userId;
            this.isSiteAdmin = isSiteAdmin;
        }

        public object Result { get; private set; }

        public List<ActionObjectPath> GetRequest(IIdProvider idProvider)
        {
            var result = new List<ActionObjectPath>();

            int webPathId = idProvider.GetActionId();
            result.Add(new ActionObjectPath
            {
                ObjectPath = new Identity { Id = webPathId, Name = webIdentity },
            });

            int siteUsersPathId = idProvider.GetActionId();
            result.Add(new ActionObjectPath
            {
                ObjectPath = new Property { Id = siteUsersPathId, ParentId = webPathId, Name = "SiteUsers" },
            });

            int userPathId = idProvider.GetActionId();
            result.Add(new ActionObjectPath
            {
                ObjectPath = new ObjectPathMethod
                {
                    Id = userPathId,
                    ParentId = siteUsersPathId,
                    Name = "GetById",
                    Parameters = new MethodParameter
                    {
                        Properties = new List<Parameter>
                        {
                            new Parameter { Type = "Int32", Value = userId },
                        },
                    },
                },
            });

            result.Add(new ActionObjectPath
            {
                Action = new SetPropertyAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = userPathId.ToString(),
                    Name = "IsSiteAdmin",
                    SetParameter = new Parameter { Type = "Boolean", Value = isSiteAdmin },
                },
            });

            result.Add(new ActionObjectPath
            {
                Action = new MethodAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = userPathId.ToString(),
                    Name = "Update",
                    Parameters = new List<Parameter>(),
                },
            });

            return result;
        }

        public void ProcessResponse(string response)
        {
        }
    }
}
