using PnP.Core.Provisioning.Model;
using PnP.Core.Services.Core.CSOM;
using PnP.Core.Services.Core.CSOM.QueryAction;
using PnP.Core.Services.Core.CSOM.QueryIdentities;
using PnP.Core.Services.Core.CSOM.Requests;
using PnP.Core.Services.Core.CSOM.Utils;
using System;
using System.Collections.Generic;
using System.Text.Json;

namespace PnP.Core.Provisioning.Services.Core.CSOM.Requests.Taxonomy
{    /// <summary>
    /// Sets a custom property on a term, term set or term group.
    /// </summary>
    /// <remarks>
    /// <para>Backlog <b>T10</b>, and one of spike <b>S1</b>'s "lossy" verdicts. CSOM distinguishes
    /// two kinds of custom property, and Graph does not:</para>
    /// <list type="bullet">
    /// <item><description><b>Shared</b> - <c>TermSetItem.SetCustomProperty</c>. Travels with the
    /// term wherever it is reused.</description></item>
    /// <item><description><b>Local</b> - <c>Term.SetLocalCustomProperty</c>. Belongs to this
    /// appearance of the term only, so a reused term can carry different values in different
    /// term sets.</description></item>
    /// </list>
    /// <para><c>ITermProperty</c> and <c>ITermSetProperty</c> are both bare key/value pairs with no
    /// way to express which kind is meant, which is why this goes through CSOM.</para>
    /// </remarks>
    internal sealed class SetTermCustomPropertyRequest : TaxonomyRequestBase, IRequest<object>
    {
        private readonly Guid termId;
        private readonly string key;
        private readonly string value;
        private readonly bool isLocal;

        /// <summary>
        /// Creates the request.
        /// </summary>
        /// <param name="termId">The term to set the property on</param>
        /// <param name="key">Property name</param>
        /// <param name="value">Property value</param>
        /// <param name="isLocal">
        /// <c>true</c> for a local property (this appearance of the term only), <c>false</c> for a
        /// shared one (travels with the term when reused).
        /// </param>
        internal SetTermCustomPropertyRequest(Guid termId, string key, string value, bool isLocal)
        {
            if (string.IsNullOrEmpty(key))
            {
                throw new ArgumentException("A property name is required.", nameof(key));
            }

            this.termId = termId;
            this.key = key;
            this.value = value;
            this.isLocal = isLocal;
        }

        public object Result { get; private set; }

        public List<ActionObjectPath> GetRequest(IIdProvider idProvider)
        {
            var paths = new List<ActionObjectPath>();

            int termStoreId = EmitTermStorePrologue(idProvider, paths);
            int termPathId = EmitGetTerm(idProvider, paths, termStoreId, termId);

            paths.Add(new ActionObjectPath
            {
                Action = new MethodAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = termPathId.ToString(),
                    Name = isLocal ? "SetLocalCustomProperty" : "SetCustomProperty",
                    Parameters = new List<Parameter>
                    {
                        new Parameter { Type = "String", Value = key },
                        new Parameter { Type = "String", Value = value }
                    }
                }
            });

            EmitCommitAll(idProvider, paths, termStoreId);

            return paths;
        }

        public void ProcessResponse(string response)
        {
        }
    }
}