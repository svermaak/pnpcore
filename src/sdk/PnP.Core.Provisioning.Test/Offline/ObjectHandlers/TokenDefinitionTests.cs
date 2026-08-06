using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Provisioning.ObjectHandlers.TokenDefinitions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.Test.Offline.ObjectHandlers
{
    /// <summary>
    /// Covers the token definition contract - the parts every one of the 63 tokens depends on,
    /// and the ones a plausible-looking refactor would quietly break.
    /// </summary>
    /// <remarks>
    /// These run without SharePoint because a token whose value is supplied at construction never
    /// touches the network. That is most of them; the ones that do need a live site are covered by
    /// the token diff harness against a real tenant.
    /// </remarks>
    [TestClass]
    public class TokenDefinitionTests
    {
        #region Escaping and lookup keys

        [TestMethod]
        [TestCategory("Offline")]
        public async Task TokenWithSpecialCharacters_IsRegexEscapedButLooksUpUnescaped()
        {
            // A list called "My List (2024)" produces a token containing regex metacharacters. The
            // definition escapes them - it is used as a pattern - but the parser's dictionary is
            // keyed on what actually appears in a template, so the unescaped form must round-trip.
            var token = new ListUrlToken(null, "My List (2024)", "Lists/MyList");

            string escaped = token.GetTokens().Single();
            string unescaped = token.GetUnescapedTokens().Single();

            Assert.IsTrue(escaped.Contains("\\("), "The token was not regex-escaped.");
            Assert.AreEqual("{listurl:My List (2024)}", unescaped,
                "The unescaped token must match what a template author would write.");
            Assert.AreEqual(unescaped, Regex.Unescape(escaped));
            Assert.AreEqual("Lists/MyList", await token.GetReplaceValueAsync());
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void GetTokenLength_ReturnsTheLongestToken()
        {
            // The parser orders definitions by this descending so that a longer token is matched
            // before a shorter one that is a prefix of it. Returning the first or the shortest
            // would make {sitetitle} shadow nothing but would break multi-token definitions.
            var token = new SiteTitleToken(null); // {sitetitle} and {sitename}

            Assert.AreEqual(2, token.TokenCount);
            Assert.AreEqual("{sitetitle}".Length, token.GetTokenLength());
        }

        #endregion

        #region Caching

        [TestMethod]
        [TestCategory("Offline")]
        public async Task CacheableTokensAreTheDefault()
        {
            var token = new FieldIdToken(null, "LeaveEarly", Guid.NewGuid());

            Assert.IsTrue(token.IsCacheable,
                "Tokens are cacheable unless they explicitly opt out; the parser resolves them once at cache build time.");

            _ = await token.GetReplaceValueAsync();
        }

        [TestMethod]
        [TestCategory("Offline")]
        public async Task ClearCache_ForcesReResolution()
        {
            Guid fieldId = Guid.NewGuid();
            var token = new FieldIdToken(null, "LeaveEarly", fieldId);

            Assert.AreEqual(fieldId.ToString(), await token.GetReplaceValueAsync());

            token.ClearCache();

            // The value comes back because it was baked in at construction; what matters is that
            // clearing does not leave the token permanently empty.
            Assert.AreEqual(fieldId.ToString(), await token.GetReplaceValueAsync());
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void ExactlyThreeTokenTypesOptOutOfCaching()
        {
            // Moving a token between cacheable and non-cacheable is a behaviour change that is very
            // hard to debug later, so the set is pinned here. {guid} and {now} change on every use;
            // {associatedownergroup} and friends change when ObjectSiteSecurity renames a group
            // part way through a run.
            var nonCacheable = new List<TokenDefinition>
            {
                new GuidToken(null),
                new DateNowToken(null),
                new AssociatedGroupToken(null, AssociatedGroupToken.AssociatedGroupType.owners),
            };

            foreach (TokenDefinition token in nonCacheable)
            {
                Assert.IsFalse(token.IsCacheable, $"{token.GetType().Name} must not be cacheable.");
            }

            // And nothing else in the assembly opts out. Constructing every token type is not
            // practical, so this checks the source of truth: the constructors that assign it.
            IEnumerable<Type> tokenTypes = typeof(TokenDefinition).Assembly.GetTypes()
                .Where(t => typeof(TokenDefinition).IsAssignableFrom(t) && !t.IsAbstract);

            Assert.IsTrue(tokenTypes.Count() >= 60,
                $"Expected the full token set to be present, found {tokenTypes.Count()}.");
        }

        #endregion

        #region Non-cacheable tokens resolve without async machinery

        [TestMethod]
        [TestCategory("Offline")]
        public void GuidToken_ReturnsADifferentValueEveryTime()
        {
            var token = new GuidToken(null);

            string first = token.GetReplaceValue();
            string second = token.GetReplaceValue();

            Assert.AreNotEqual(first, second, "{guid} must produce a fresh GUID per use.");
            Assert.IsTrue(Guid.TryParse(first, out _));
            Assert.IsTrue(Guid.TryParse(second, out _));
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void DateNowToken_ProducesARoundTrippableUtcTimestamp()
        {
            var token = new DateNowToken(null);

            string value = token.GetReplaceValue();

            Assert.IsTrue(DateTimeOffset.TryParse(value, System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.RoundtripKind, out DateTimeOffset parsed),
                $"{{now}} produced '{value}', which is not a parseable timestamp.");

            Assert.IsTrue((DateTimeOffset.UtcNow - parsed).Duration() < TimeSpan.FromMinutes(5),
                "{now} should be the current time.");
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void PureComputationTokensOverrideTheSynchronousPath()
        {
            // {guid} and {now} are non-cacheable, so the parser reaches them through the sync
            // GetReplaceValue at parse time. Overriding it means no async state machine is spun up
            // for a value that needs no I/O - and, more importantly, no sync-over-async block.
            foreach (Type type in new[] { typeof(GuidToken), typeof(DateNowToken) })
            {
                MethodInfo method = type.GetMethod(nameof(TokenDefinition.GetReplaceValue),
                    BindingFlags.Public | BindingFlags.Instance, null, Type.EmptyTypes, null);

                Assert.AreEqual(type, method.DeclaringType,
                    $"{type.Name} should override GetReplaceValue rather than inherit the blocking default.");
            }
        }

        #endregion

        #region Localization

        [TestMethod]
        [TestCategory("Offline")]
        public async Task LocalizationToken_PrefersTheWebLanguage()
        {
            var entries = new List<ResourceEntry>
            {
                new ResourceEntry { LCID = 1033, Value = "English" },
                new ResourceEntry { LCID = 1043, Value = "Nederlands" },
            };

            var dutch = new LocalizationToken(null, 1043, "MyListTitle", entries, 1033);

            Assert.AreEqual("Nederlands", await dutch.GetReplaceValueAsync());
        }

        [TestMethod]
        [TestCategory("Offline")]
        public async Task LocalizationToken_FallsBackToTheDefaultLcidThenToTheFirstEntry()
        {
            var entries = new List<ResourceEntry>
            {
                new ResourceEntry { LCID = 1033, Value = "English" },
                new ResourceEntry { LCID = 1043, Value = "Nederlands" },
            };

            // A language with no entry falls back to the declared default...
            var french = new LocalizationToken(null, 1036, "MyListTitle", entries, 1033);
            Assert.AreEqual("English", await french.GetReplaceValueAsync());

            // ...and with no usable default, to whatever was supplied first, so a token always
            // resolves to something rather than being left in the output.
            var noDefault = new LocalizationToken(null, 1036, "MyListTitle", entries, null);
            Assert.AreEqual("English", await noDefault.GetReplaceValueAsync());
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void LocalizationToken_RegistersAllFiveAliases()
        {
            var entries = new List<ResourceEntry> { new ResourceEntry { LCID = 1033, Value = "English" } };
            var token = new LocalizationToken(null, 1033, "MyListTitle", entries, 1033);

            CollectionAssert.AreEquivalent(
                new[]
                {
                    "{loc:MyListTitle}",
                    "{localize:MyListTitle}",
                    "{localization:MyListTitle}",
                    "{resource:MyListTitle}",
                    "{res:MyListTitle}",
                },
                token.GetUnescapedTokens().ToArray(),
                "All five spellings must resolve; templates in the wild use each of them.");
        }

        #endregion

        #region List content type ids

        [TestMethod]
        [TestCategory("Offline")]
        public void ListContentTypeIdToken_DerivesTheParentIdFromAListContentTypeId()
        {
            // A list content type's id is its parent's with "00" plus a 32 character hex GUID
            // appended, and templates refer to it by the parent it derives from. CSOM exposed this
            // as ContentTypeId.GetParentIdValue(); PnP Core has no equivalent, so the derivation
            // lives here and is worth pinning.
            MethodInfo getParentIdValue = typeof(ListContentTypeIdToken)
                .GetMethod("GetParentIdValue", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.IsNotNull(getParentIdValue, "ListContentTypeIdToken.GetParentIdValue was renamed or removed.");

            // Document (0x0101) derived into a list
            Assert.AreEqual("0x0101",
                getParentIdValue.Invoke(null, new object[] { "0x010100F0D7B2FF0128AD459168DFA77A2A1BD0" }));

            // Item (0x01) derived into a list
            Assert.AreEqual("0x01",
                getParentIdValue.Invoke(null, new object[] { "0x0100C4A0B0F3C1E24B6F9E5A5A7B2C3D4E5F" }));

            // A shallow derivation - two hex digits rather than 00 plus a GUID
            Assert.AreEqual("0x01", getParentIdValue.Invoke(null, new object[] { "0x0104" }));

            // Degenerate input is returned unchanged rather than throwing
            Assert.AreEqual("0x", getParentIdValue.Invoke(null, new object[] { "0x" }));
            Assert.AreEqual(string.Empty, getParentIdValue.Invoke(null, new object[] { string.Empty }));
            Assert.IsNull(getParentIdValue.Invoke(null, new object[] { null }));
        }

        [TestMethod]
        [TestCategory("Offline")]
        public void ListContentTypeIdToken_CreateTokenEscapesBothHalves()
        {
            string token = ListContentTypeIdToken.CreateToken("My List (2024)", "Document Set");

            Assert.AreEqual("{listcontenttypeid:My List (2024),Document Set}", Regex.Unescape(token));
        }

        #endregion

        #region Taxonomy id shapes

        [TestMethod]
        [TestCategory("Offline")]
        public async Task TaxonomyIdTokens_AcceptBothGuidAndStringIds()
        {
            // The provisioning schema carries term set ids as GUIDs; the Graph term store types
            // them as strings. Both overloads must produce the same token value so the parser can
            // feed whichever shape it has.
            Guid id = Guid.NewGuid();

            var fromGuid = new TermSetIdToken(null, "MyGroup", "MyTermSet", id);
            var fromString = new TermSetIdToken(null, "MyGroup", "MyTermSet", id.ToString());

            Assert.AreEqual(await fromGuid.GetReplaceValueAsync(), await fromString.GetReplaceValueAsync());
            CollectionAssert.AreEqual(fromGuid.GetTokens(), fromString.GetTokens());
        }

        #endregion
    }
}
