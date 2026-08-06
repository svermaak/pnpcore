using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.ObjectHandlers.TokenDefinitions
{
    /// <summary>
    /// Defines a provisioning engine token - one or more <c>{placeholder}</c> strings and the
    /// value they resolve to.
    /// </summary>
    /// <remarks>
    /// <para><b>Async resolution, sync parsing.</b> PnP Framework's <c>GetReplaceValue()</c> was
    /// synchronous because CSOM is. PnP Core is async throughout, so the primitive a token
    /// implements is <see cref="GetReplaceValueAsync"/>. Parsing stays synchronous, which is what
    /// makes the handler ports tractable - <c>ParseString</c> is called from hundreds of places,
    /// including inside LINQ expressions that cannot await.</para>
    /// <para>That combination works because the parser resolves eagerly: for a <em>cacheable</em>
    /// token (the default, and all but three of the 63) the value is awaited once while the token
    /// cache is built, and parsing only ever reads a dictionary. Only a non-cacheable token can
    /// reach the sync path at parse time. See <see cref="GetReplaceValue"/>.</para>
    /// <para><b>No cloned context.</b> PnP Framework gave every token a <c>TokenContext</c> - a
    /// cloned <c>ClientContext</c> - so a token's <c>Load</c> could not piggyback on the caller's
    /// pending <c>ExecuteQuery</c> batch, and the parser threw if a token dirtied the current
    /// context. PnP Core has no such hazard: an un-batched call executes on its own and returns.
    /// Tokens therefore use <see cref="Context"/> directly, and both the clone and the guard are
    /// gone.</para>
    /// </remarks>
    public abstract class TokenDefinition
    {
        /// <summary>
        /// The resolved value, once known. Set by <see cref="GetReplaceValueAsync"/> implementations
        /// and cleared by <see cref="ClearCache"/>.
        /// </summary>
        protected string CacheValue;

        private readonly string[] _tokens;
        private readonly string[] _unescapedTokens;
        private readonly int _maximumTokenLength;

        /// <summary>
        /// Whether this token's value may be resolved once and reused. Defaults to <c>true</c>.
        /// </summary>
        /// <remarks>
        /// A cacheable token is resolved while the parser builds its cache and never resolved
        /// again. Set this to <c>false</c> only when the value can legitimately change during a
        /// provisioning run - a fresh GUID per use, the current time, or a name a later handler
        /// rewrites. Moving a token between the two is a behaviour change and is hard to debug
        /// after the fact.
        /// </remarks>
        public bool IsCacheable { get; set; } = true;

        /// <summary>
        /// Creates a token definition bound to a context.
        /// </summary>
        /// <param name="context">The context the token resolves against. May be <c>null</c> for
        /// tokens whose value is supplied at construction and needs no lookup.</param>
        /// <param name="token">One or more token strings, already <see cref="Regex"/>-escaped
        /// where the token embeds caller-supplied text.</param>
        protected TokenDefinition(PnPContext context, params string[] token)
        {
            Context = context;
            _tokens = token;
            _unescapedTokens = GetUnescapedTokens(token);
            _maximumTokenLength = GetMaximumTokenLength(token);
        }

        /// <summary>
        /// The context this token resolves against.
        /// </summary>
        public PnPContext Context { get; set; }

        /// <summary>
        /// The number of token strings this definition supplies.
        /// </summary>
        public int TokenCount => _tokens.Length;

        /// <summary>
        /// The token strings, as constructed (i.e. still regex-escaped).
        /// </summary>
        public string[] GetTokens()
        {
            return _tokens;
        }

        /// <summary>
        /// The token strings with <see cref="Regex.Unescape(string)"/> applied - the form that
        /// actually appears in a template, and therefore the form used as a dictionary key.
        /// </summary>
        public IReadOnlyList<string> GetUnescapedTokens()
        {
            return _unescapedTokens;
        }

        /// <summary>
        /// The length of the longest token string. The parser orders definitions by this
        /// descending, so that <c>{listid:Foo}</c> is considered before a shorter token that
        /// happens to be a prefix of it.
        /// </summary>
        public int GetTokenLength()
        {
            return _maximumTokenLength;
        }

        /// <summary>
        /// Resolves the token's value.
        /// </summary>
        public abstract Task<string> GetReplaceValueAsync();

        /// <summary>
        /// Resolves the token's value synchronously.
        /// </summary>
        /// <remarks>
        /// Only reached for non-cacheable tokens, since cacheable ones are resolved and stored
        /// while the cache is built. Tokens that compute their value without I/O - <c>{guid}</c>
        /// and <c>{now}</c> - override this so no async machinery is involved at all.
        /// </remarks>
        public virtual string GetReplaceValue()
        {
            return GetReplaceValueAsync().GetAwaiter().GetResult();
        }

        /// <summary>
        /// Discards the resolved value so the next resolution goes back to the source.
        /// </summary>
        public void ClearCache()
        {
            CacheValue = null;
        }

        private static int GetMaximumTokenLength(IReadOnlyList<string> tokens)
        {
            var result = 0;

            for (var index = 0; index < tokens.Count; index++)
            {
                result = Math.Max(result, tokens[index].Length);
            }

            return result;
        }

        private static string[] GetUnescapedTokens(IReadOnlyList<string> tokens)
        {
            var result = new string[tokens.Count];

            for (var index = 0; index < tokens.Count; index++)
            {
                result[index] = Regex.Unescape(tokens[index]);
            }

            return result;
        }
    }
}
