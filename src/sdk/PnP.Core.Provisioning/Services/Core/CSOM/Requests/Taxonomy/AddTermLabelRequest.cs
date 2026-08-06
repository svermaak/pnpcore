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
    /// Adds a label to a term in a given language.
    /// </summary>
    /// <remarks>
    /// Backlog <b>T10</b>. <c>Term.CreateLabel(value, lcid, isDefaultForLanguage)</c>. A term's
    /// labels are how it is spelled in each language plus any synonyms; a template's
    /// <c>&lt;pnp:Label&gt;</c> elements map one-to-one onto these.
    /// </remarks>
    internal sealed class AddTermLabelRequest : TaxonomyRequestBase, IRequest<object>
    {
        private readonly Guid termId;
        private readonly string value;
        private readonly int lcid;
        private readonly bool isDefaultForLanguage;

        internal AddTermLabelRequest(Guid termId, string value, int lcid, bool isDefaultForLanguage)
        {
            if (string.IsNullOrEmpty(value))
            {
                throw new ArgumentException("A label value is required.", nameof(value));
            }

            this.termId = termId;
            this.value = value;
            this.lcid = lcid;
            this.isDefaultForLanguage = isDefaultForLanguage;
        }

        public object Result { get; private set; }

        public List<ActionObjectPath> GetRequest(IIdProvider idProvider)
        {
            var paths = new List<ActionObjectPath>();

            int termStoreId = EmitTermStorePrologue(idProvider, paths);
            int termPathId = EmitGetTerm(idProvider, paths, termStoreId, termId);

            EmitMethod(idProvider, paths, termPathId, "CreateLabel",
                new Parameter { Type = "String", Value = value },
                new Parameter { Type = "Number", Value = lcid },
                new Parameter { Type = "Boolean", Value = isDefaultForLanguage });

            EmitCommitAll(idProvider, paths, termStoreId);

            return paths;
        }

        public void ProcessResponse(string response)
        {
        }
    }
}