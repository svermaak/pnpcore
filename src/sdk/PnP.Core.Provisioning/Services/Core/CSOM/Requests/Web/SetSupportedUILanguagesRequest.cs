using PnP.Core.Services.Core.CSOM;
using PnP.Core.Services.Core.CSOM.QueryAction;
using PnP.Core.Services.Core.CSOM.QueryIdentities;
using PnP.Core.Services.Core.CSOM.Requests;
using PnP.Core.Services.Core.CSOM.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Core.Provisioning.Services.Core.CSOM.Requests.Web
{
    /// <summary>
    /// Adds and removes the web's supported UI languages in one round trip.
    /// </summary>
    /// <remarks>
    /// <para>Backlog <b>T2</b>. This exists because <b>REST silently does nothing</b>:
    /// <c>_api/web/AddSupportedUILanguage(lcid=…)</c> returns success and has no effect. That was
    /// diagnosed rather than assumed — the language pack was installed and the web was already
    /// multilingual, so neither of the other candidate explanations held.</para>
    /// <para><b>Why REST fails is now obvious in hindsight:</b> <c>AddSupportedUILanguage</c> only
    /// <em>stages</em> the change, and <c>Web.Update()</c> is what persists it. The REST endpoint
    /// offers no way to express that second step. This is the <b>fifth</b> staged-write trap in the
    /// migration, after the term store's <c>CommitAll()</c>, <c>Audit.Update()</c>,
    /// <c>SaveWebPartChanges()</c> and the user resource parent update.</para>
    /// <para><b>Removals go before additions</b>, matching PnP Framework. A template is the intended
    /// final state, and doing it the other way round can transiently leave the web with no
    /// languages at all.</para>
    /// </remarks>
    internal sealed class SetSupportedUILanguagesRequest : IRequest<object>
    {
        private readonly Guid siteId;
        private readonly Guid webId;
        private readonly IReadOnlyCollection<int> toAdd;
        private readonly IReadOnlyCollection<int> toRemove;

        internal SetSupportedUILanguagesRequest(Guid siteId, Guid webId,
            IReadOnlyCollection<int> toAdd, IReadOnlyCollection<int> toRemove)
        {
            this.siteId = siteId;
            this.webId = webId;
            this.toAdd = toAdd ?? Array.Empty<int>();
            this.toRemove = toRemove ?? Array.Empty<int>();

            if (this.toAdd.Count == 0 && this.toRemove.Count == 0)
            {
                throw new ArgumentException("At least one language to add or remove is required.", nameof(toAdd));
            }
        }

        public object Result { get; private set; }

        /// <summary>
        /// The raw CSOM response body.
        /// </summary>
        /// <remarks>
        /// Captured because <c>ApiCallResponse.Json</c> is not populated on the CSOM path - the only
        /// place <c>ErrorInfo</c> is visible is here. Every CSOM request in this project should keep
        /// this; it is what turned the regional settings investigation from guesswork into evidence.
        /// </remarks>
        internal string RawResponse { get; private set; }

        public List<ActionObjectPath> GetRequest(IIdProvider idProvider)
        {
            var paths = new List<ActionObjectPath>();

            int webIdentityId = idProvider.GetActionId();

            paths.Add(new ActionObjectPath
            {
                ObjectPath = new Identity
                {
                    Id = webIdentityId,
                    Name = CsomIdentity.Web(siteId, webId)
                }
            });

            // Remove first - see the remarks.
            foreach (int lcid in toRemove)
            {
                paths.Add(LanguageMethod(idProvider, webIdentityId, "RemoveSupportedUILanguage", lcid));
            }

            foreach (int lcid in toAdd)
            {
                paths.Add(LanguageMethod(idProvider, webIdentityId, "AddSupportedUILanguage", lcid));
            }

            // Persist. Without this the adds and removes are staged and discarded - which is
            // precisely how the REST attempt failed.
            paths.Add(new ActionObjectPath
            {
                Action = new MethodAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = webIdentityId.ToString(),
                    Name = "Update",
                    Parameters = new List<Parameter>()
                }
            });

            return paths;
        }

        private static ActionObjectPath LanguageMethod(IIdProvider idProvider, int webIdentityId, string name, int lcid)
        {
            return new ActionObjectPath
            {
                Action = new MethodAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = webIdentityId.ToString(),
                    Name = name,
                    Parameters = new List<Parameter>
                    {
                        new Parameter { Type = "Number", Value = lcid }
                    }
                }
            };
        }

        public void ProcessResponse(string response)
        {
            RawResponse = response;
        }
    }
}
