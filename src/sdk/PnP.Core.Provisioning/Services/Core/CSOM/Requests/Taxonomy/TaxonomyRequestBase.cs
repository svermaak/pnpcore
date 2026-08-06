using PnP.Core.Services.Core.CSOM;
using PnP.Core.Services.Core.CSOM.QueryAction;
using PnP.Core.Services.Core.CSOM.QueryIdentities;
using PnP.Core.Services.Core.CSOM.Utils;
using System;
using System.Collections.Generic;

namespace PnP.Core.Provisioning.Services.Core.CSOM.Requests.Taxonomy
{
    /// <summary>
    /// Shared plumbing for the taxonomy CSOM requests: the object path prologue that every one of
    /// them starts with, and the term store navigation they build on top of it.
    /// </summary>
    /// <remarks>
    /// <para>These requests exist because of spike <b>S1</b>. PnP Core's Graph term store cannot
    /// preserve a caller-supplied term GUID - every <c>Id</c> is read-only and no <c>AddAsync</c>
    /// overload accepts one - and it has no representation at all for custom sort order,
    /// deprecation, or the shared-versus-local custom property distinction. Provisioning templates
    /// depend on all four, so taxonomy provisioning goes through CSOM. Decision <b>D7</b>.</para>
    /// <para>Every taxonomy call starts the same way: get the taxonomy session, then the default
    /// site collection term store. That prologue is three object paths and six actions, so it lives
    /// here rather than being copy-pasted eight times.</para>
    /// </remarks>
    internal abstract class TaxonomyRequestBase
    {
        internal CSOMResponseHelper ResponseHelper { get; set; } = new CSOMResponseHelper();

        /// <summary>
        /// Emits <c>TaxonomySession.GetTaxonomySession()</c> followed by
        /// <c>GetDefaultSiteCollectionTermStore()</c>, and returns the term store's object path id.
        /// </summary>
        /// <param name="idProvider">The request's id provider</param>
        /// <param name="paths">The action list being built; the prologue is appended to it</param>
        /// <returns>The object path id of the term store, to hang further calls off</returns>
        protected static int EmitTermStorePrologue(IIdProvider idProvider, List<ActionObjectPath> paths)
        {
            var session = new StaticMethodPath
            {
                Id = idProvider.GetActionId(),
                TypeId = CsomTypeIds.TaxonomySession,
                Name = "GetTaxonomySession",
                Parameters = new MethodParameter { Properties = new List<Parameter>() }
            };

            paths.Add(new ActionObjectPath
            {
                Action = new BaseAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = session.Id.ToString()
                },
                ObjectPath = session
            });

            paths.Add(new ActionObjectPath
            {
                Action = new IdentityQueryAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = session.Id.ToString()
                }
            });

            var termStore = new ObjectPathMethod
            {
                Id = idProvider.GetActionId(),
                ParentId = session.Id,
                Name = "GetDefaultSiteCollectionTermStore",
                Parameters = new MethodParameter { Properties = new List<Parameter>() }
            };

            paths.Add(new ActionObjectPath
            {
                Action = new BaseAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = termStore.Id.ToString()
                },
                ObjectPath = termStore
            });

            paths.Add(new ActionObjectPath
            {
                Action = new IdentityQueryAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = termStore.Id.ToString()
                }
            });

            return termStore.Id;
        }

        /// <summary>
        /// Emits <c>TermStore.GetGroup(id)</c> and returns the group's object path id.
        /// </summary>
        protected static int EmitGetGroup(IIdProvider idProvider, List<ActionObjectPath> paths, int termStorePathId, Guid groupId)
        {
            return EmitMethod(idProvider, paths, termStorePathId, "GetGroup",
                new Parameter { Type = "Guid", Value = groupId });
        }

        /// <summary>
        /// Emits <c>TermStore.GetTermSet(id)</c> and returns the term set's object path id.
        /// </summary>
        protected static int EmitGetTermSet(IIdProvider idProvider, List<ActionObjectPath> paths, int termStorePathId, Guid termSetId)
        {
            return EmitMethod(idProvider, paths, termStorePathId, "GetTermSet",
                new Parameter { Type = "Guid", Value = termSetId });
        }

        /// <summary>
        /// Emits <c>TermStore.GetTerm(id)</c> and returns the term's object path id.
        /// </summary>
        protected static int EmitGetTerm(IIdProvider idProvider, List<ActionObjectPath> paths, int termStorePathId, Guid termId)
        {
            return EmitMethod(idProvider, paths, termStorePathId, "GetTerm",
                new Parameter { Type = "Guid", Value = termId });
        }

        /// <summary>
        /// Emits a method call that returns a client object, with an identity query so the result
        /// can be referenced by later actions.
        /// </summary>
        protected static int EmitMethod(IIdProvider idProvider, List<ActionObjectPath> paths, int parentPathId, string name, params Parameter[] parameters)
        {
            var method = new ObjectPathMethod
            {
                Id = idProvider.GetActionId(),
                ParentId = parentPathId,
                Name = name,
                Parameters = new MethodParameter { Properties = new List<Parameter>(parameters) }
            };

            paths.Add(new ActionObjectPath
            {
                Action = new BaseAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = method.Id.ToString()
                },
                ObjectPath = method
            });

            paths.Add(new ActionObjectPath
            {
                Action = new IdentityQueryAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = method.Id.ToString()
                }
            });

            return method.Id;
        }

        /// <summary>
        /// Emits <c>TermStore.CommitAll()</c>.
        /// </summary>
        /// <remarks>
        /// Taxonomy writes are staged and do not take effect until the store is committed. Omitting
        /// this makes a request appear to succeed while changing nothing - the single most common
        /// way to get CSOM taxonomy wrong.
        /// </remarks>
        protected static void EmitCommitAll(IIdProvider idProvider, List<ActionObjectPath> paths, int termStorePathId)
        {
            paths.Add(new ActionObjectPath
            {
                Action = new MethodAction
                {
                    Id = idProvider.GetActionId(),
                    ObjectPathId = termStorePathId.ToString(),
                    Name = "CommitAll",
                    Parameters = new List<Parameter>()
                }
            });
        }
    }
}
