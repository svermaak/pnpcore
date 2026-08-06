using PnP.Core.Services;

namespace PnP.Core.Provisioning.ObjectHandlers.TokenDefinitions
{
    /// <summary>
    /// A token whose value is tied to a particular web, and so must be re-resolved when the
    /// engine moves to a different one.
    /// </summary>
    /// <remarks>
    /// Applying a hierarchy walks several sites with one parser. <c>RebaseAsync</c> calls
    /// <see cref="ClearVolatileCache"/> on every volatile token, repointing it at the new context
    /// and discarding the old site's value. A non-volatile token keeps its value across the move -
    /// which is correct for e.g. <c>{parameter:...}</c> and wrong for <c>{site}</c>.
    /// </remarks>
    public abstract class VolatileTokenDefinition : TokenDefinition
    {
        /// <summary>
        /// Creates a volatile token definition.
        /// </summary>
        /// <param name="context">The context the token resolves against</param>
        /// <param name="token">One or more token strings</param>
        protected VolatileTokenDefinition(PnPContext context, params string[] token) : base(context, token)
        {
        }

        /// <summary>
        /// Repoints the token at a new context and discards the value resolved against the old one.
        /// </summary>
        /// <param name="context">The context to resolve against from now on</param>
        public void ClearVolatileCache(PnPContext context)
        {
            CacheValue = null;
            Context = context;
        }
    }
}
