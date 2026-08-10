using Microsoft.Extensions.Logging;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Services.Core.CSOM;
using PnP.Core.Provisioning.Services.Core.CSOM.Requests;
using PnP.Core.Provisioning.Services.Core.CSOM.Requests.UserResources;
using PnP.Core.Services;
using PnP.Core.Services.Core.CSOM.Requests;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.ObjectHandlers.Utilities
{
    /// <summary>
    /// Reads and writes the per-language values behind a <c>{res:…}</c> token.
    /// </summary>
    /// <remarks>
    /// <para><b>Backlog T6, and the reason phase 6 starts here.</b> A dozen handlers carry a
    /// <c>Title</c> or <c>Description</c> that a template can localize - fields, content types,
    /// lists, views, custom actions, navigation nodes. Landing this before the wave-2 handlers means
    /// they are written <em>with</em> the resource path rather than retrofitted, and it is where the
    /// phase-5 handlers get the coverage they deferred.</para>
    /// <para><b>PnP Core cannot do this yet.</b> <c>IUserResource</c> is commented out across
    /// <c>IContentType</c>, <c>IField</c>, <c>IUserCustomAction</c> and <c>IWeb</c> - a known
    /// deferral upstream - so this reaches <c>SP.UserResource</c> through the CSOM requests built in
    /// phase 4. When PnP Core models it, this class becomes a forwarder and no handler changes.</para>
    /// </remarks>
    internal static class UserResources
    {
        /// <summary>
        /// Matches every spelling of a localization token the schema allows.
        /// </summary>
        private static readonly Regex ResourceTokenPattern =
            new Regex("\\{(res|loc|resource|localize|localization):(.*?)(\\})", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Whether a value is a localization token rather than a literal.
        /// </summary>
        internal static bool ContainsResourceToken(string value)
        {
            return value != null && ResourceTokenPattern.IsMatch(value);
        }

        #region Apply

        /// <summary>
        /// Writes every language a token resolves to onto one resource property.
        /// </summary>
        /// <param name="context">The site being provisioned</param>
        /// <param name="resource">Which resource property on which object</param>
        /// <param name="tokenValue">The <c>{res:…}</c> token from the template</param>
        /// <param name="parser">The parser holding the template's resource files</param>
        /// <returns>Whether anything was written</returns>
        /// <remarks>
        /// <para>All cultures go in <b>one</b> CSOM batch. A template that localizes into eight
        /// languages should not cost eight round trips, and the round trip dominates the cost.</para>
        /// <para><b>Each culture carries its own parent <c>Update</c>.</b> That looks redundant in a
        /// batch and is not: <c>SetValueForUICulture</c> only stages the change, and the update has
        /// to follow it. Collapsing them into a single trailing update was tried and is what made the
        /// first live run of these requests silently write nothing.</para>
        /// </remarks>
        internal static async Task<bool> SetAsync(PnPContext context, UserResourcePath resource, string tokenValue, TokenParser parser)
        {
            if (resource == null || string.IsNullOrEmpty(tokenValue) || !ContainsResourceToken(tokenValue))
            {
                return false;
            }

            List<Tuple<string, string>> resourceValues = parser.GetResourceTokenResourceValues(tokenValue);
            if (resourceValues.Count == 0)
            {
                context.Logger?.LogWarning(
                    "{Source}: '{Token}' resolved to no languages. Is the template's resource file listed under <pnp:Localizations>?",
                    Constants.LOGGING_SOURCE, tokenValue);

                return false;
            }

            var requests = resourceValues
                .Select(value => (IRequest<object>)new SetValueForUICultureRequest(resource, value.Item1, value.Item2))
                .ToList();

            await CsomRequestSender.SendManyAsync(context, requests).ConfigureAwait(false);

            context.Logger?.LogInformation("{Source}: localized {Property} into {Count} language(s).",
                Constants.LOGGING_SOURCE, resource.PropertyName, requests.Count);

            return true;
        }

        /// <summary>
        /// Writes a token's languages onto a resource property, reporting failure rather than
        /// throwing.
        /// </summary>
        /// <remarks>
        /// Localization is decoration on an object that has already been provisioned correctly.
        /// Failing the whole run because one title could not be translated trades a complete site
        /// for a cosmetic problem, so handlers that are mid-way through their work use this.
        /// </remarks>
        internal static async Task<bool> TrySetAsync(PnPContext context, UserResourcePath resource, string tokenValue,
            TokenParser parser, string description, Action<string> reportWarning = null)
        {
            try
            {
                return await SetAsync(context, resource, tokenValue, parser).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                string message = string.Format(CultureInfo.CurrentCulture,
                    "Could not localize {0}: {1}", description, ex.Message);

                context.Logger?.LogWarning(ex, "{Source}: {Message}", Constants.LOGGING_SOURCE, message);
                reportWarning?.Invoke(message);

                return false;
            }
        }

        #endregion

        #region Extract

        /// <summary>
        /// Reads a resource property in every language the template supports and records the values
        /// against a token.
        /// </summary>
        /// <returns>
        /// Whether any language had a value - which is what tells the caller to replace the literal
        /// title with the token.
        /// </returns>
        /// <remarks>
        /// A property with no translations returns nothing, and the caller must then keep the plain
        /// value. Writing the token regardless would produce a template whose titles resolve to
        /// nothing on the target.
        /// </remarks>
        internal static async Task<bool> PersistAsync(PnPContext context, UserResourcePath resource, string token,
            ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInformation)
        {
            if (resource == null || template.SupportedUILanguages.Count == 0)
            {
                return false;
            }

            // All languages in ONE round trip.
            //
            // This is not a micro-optimisation. A real tenant readily has 50+ supported UI languages,
            // and wave 2 localizes two properties on every field, content type, list and view - so
            // one request per language would be thousands of round trips for a single extract.
            // IRequest<T> is covariant, so the typed read requests batch without an adapter.
            var requests = new List<(int Lcid, GetValueForUICultureRequest Request)>();

            foreach (SupportedUILanguage language in template.SupportedUILanguages)
            {
                requests.Add((language.LCID, new GetValueForUICultureRequest(resource, new CultureInfo(language.LCID).Name)));
            }

            try
            {
                await CsomRequestSender.SendManyAsync(context,
                    requests.Select(r => (IRequest<object>)r.Request).ToList()).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                context.Logger?.LogDebug(ex, "{Source}: could not read {Property} in any language.",
                    Constants.LOGGING_SOURCE, resource.PropertyName);
                return false;
            }

            bool any = false;

            foreach ((int lcid, GetValueForUICultureRequest request) in requests)
            {
                if (string.IsNullOrEmpty(request.Result))
                {
                    continue;
                }

                any = true;
                creationInformation.ResourceTokens.Add(new Tuple<string, int, string>(token, lcid, request.Result));
            }

            return any;
        }

        /// <summary>
        /// Reads a localizable property in every language a set of LCIDs names, in one round trip.
        /// </summary>
        /// <returns>
        /// The value per LCID, with languages that have no translation left out.
        /// </returns>
        /// <remarks>
        /// Separate from <see cref="PersistAsync"/> because not every reader is persisting into a
        /// template: <c>ObjectListInstance</c> reads a list's translated titles purely to register
        /// a <c>{listid:…}</c> token for each of them, so a template written in one language still
        /// resolves against a site rendered in another.
        /// </remarks>
        internal static async Task<Dictionary<int, string>> ReadAsync(PnPContext context, UserResourcePath resource,
            IEnumerable<int> lcids)
        {
            var values = new Dictionary<int, string>();

            if (resource == null || lcids == null)
            {
                return values;
            }

            var requests = lcids.Distinct()
                .Select(lcid => (Lcid: lcid, Request: new GetValueForUICultureRequest(resource, new CultureInfo(lcid).Name)))
                .ToList();

            if (requests.Count == 0)
            {
                return values;
            }

            await CsomRequestSender.SendManyAsync(context,
                requests.Select(r => (IRequest<object>)r.Request).ToList()).ConfigureAwait(false);

            foreach ((int lcid, GetValueForUICultureRequest request) in requests)
            {
                if (!string.IsNullOrEmpty(request.Result))
                {
                    values[lcid] = request.Result;
                }
            }

            return values;
        }

        /// <summary>
        /// Records a value that was already read, against a token and a language.
        /// </summary>
        /// <remarks>
        /// Used where the value does not come from a <c>UserResource</c> at all - a view title read
        /// through an <c>Accept-Language</c> request, or a footer node title. Same resource file,
        /// different source.
        /// </remarks>
        internal static bool PersistValue(string token, int lcid, string value, ProvisioningTemplateCreationInformation creationInformation)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            creationInformation.ResourceTokens.Add(new Tuple<string, int, string>(token, lcid, value));

            return true;
        }

        /// <summary>
        /// The token a caller should write into the template once persisting succeeded.
        /// </summary>
        internal static string TokenFor(string key)
        {
            return $"{{res:{key}}}";
        }

        #endregion
    }
}
