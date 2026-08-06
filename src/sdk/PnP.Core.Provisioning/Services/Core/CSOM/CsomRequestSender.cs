using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using PnP.Core.Services.Core;
using PnP.Core.Services.Core.CSOM.Requests;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.Services.Core.CSOM
{
    /// <summary>
    /// Sends CSOM requests through PnP Core's batching pipeline.
    /// </summary>
    /// <remarks>
    /// <para>PnP Core executes a CSOM request by wrapping it in an <see cref="ApiCall"/> and
    /// posting it through the internal <c>Web</c> model's <c>RawRequestAsync</c>. That is three
    /// lines of ceremony plus two casts to internal types, repeated at every call site - so it
    /// lives here once instead of thirty times.</para>
    /// <para>Reachable at all because of decision <b>D3</b>: <c>PnP.Core.Provisioning</c> is a
    /// friend assembly, so <c>IRequest&lt;T&gt;</c>, <c>ApiCall</c> and the internal <c>Web</c>
    /// are all in scope.</para>
    /// </remarks>
    internal static class CsomRequestSender
    {
        /// <summary>
        /// Sends a single CSOM request and returns its typed result.
        /// </summary>
        /// <typeparam name="T">
        /// The request's result type. <b>Must be a reference type</b> - see the remarks.
        /// </typeparam>
        /// <param name="context">The context to send against</param>
        /// <param name="request">The request to send</param>
        /// <returns>The request's <c>Result</c> after the response has been processed</returns>
        /// <remarks>
        /// <para><b>The result type must be a reference type.</b> Batching casts every request to
        /// <c>IRequest&lt;object&gt;</c>, and <c>IRequest&lt;out T&gt;</c>'s covariance does not
        /// apply to value types. A request declared as, say, <c>IRequest&lt;int&gt;</c> or
        /// <c>IRequest&lt;SomeEnum?&gt;</c> compiles cleanly and then throws
        /// <see cref="InvalidCastException"/> at run time from inside the SDK.</para>
        /// <para>The generic parameter cannot carry a <c>where T : class</c> constraint without
        /// diverging from <c>IRequest&lt;T&gt;</c>, so this is enforced by a test that reflects
        /// over every request type in the assembly instead.</para>
        /// </remarks>
        internal static async Task<T> SendAsync<T>(PnPContext context, IRequest<T> request)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            await SendManyAsync(context, new List<IRequest<object>> { (IRequest<object>)request }).ConfigureAwait(false);

            return request.Result;
        }

        /// <summary>
        /// Sends several CSOM requests in a single round trip.
        /// </summary>
        /// <remarks>
        /// CSOM's own batching: the requests share one <c>ProcessQuery</c> POST, and each one's
        /// <c>ProcessResponse</c> is called with the shared response body. Worth using wherever a
        /// handler has independent work to do - provisioning a term set's labels, for instance -
        /// since the round trip dominates the cost.
        /// </remarks>
        /// <param name="context">The context to send against</param>
        /// <param name="requests">The requests to send, in order</param>
        internal static async Task SendManyAsync(PnPContext context, List<IRequest<object>> requests)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            if (requests == null || requests.Count == 0)
            {
                throw new ArgumentException("At least one request is required.", nameof(requests));
            }

            var apiCall = new ApiCall(requests);

            await (context.Web as Web).RawRequestAsync(apiCall, HttpMethod.Post).ConfigureAwait(false);
        }

        /// <summary>
        /// Loads the site and web ids, which almost every CSOM identity string needs.
        /// </summary>
        /// <remarks>
        /// Both are cheap and almost always already loaded by the time a handler runs, so this is a
        /// no-op in the common case rather than a round trip.
        /// </remarks>
        internal static async Task<(Guid SiteId, Guid WebId)> GetSiteAndWebIdAsync(PnPContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            await context.Site.EnsurePropertiesAsync(s => s.Id).ConfigureAwait(false);
            await context.Web.EnsurePropertiesAsync(w => w.Id).ConfigureAwait(false);

            return (context.Site.Id, context.Web.Id);
        }
    }
}
