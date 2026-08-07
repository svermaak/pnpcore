using Microsoft.Extensions.Logging;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.Provisioning.ObjectHandlers.Utilities;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using FooterModel = PnP.Core.Provisioning.Model.SiteFooter;

namespace PnP.Core.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Reads and writes the modern site footer: its chrome (enabled, layout, emphasis, display name)
    /// and its navigation - the logo, the title and the link tree.
    /// </summary>
    /// <remarks>
    /// <para><b>The footer is two different things behind one template element.</b> The chrome half
    /// is <see cref="IBrandingManager"/>; the navigation half is a hidden navigation tree that only
    /// the <c>MenuState</c> / <c>SaveMenuState</c> endpoints reach, keyed by four fixed GUIDs (see
    /// <c>Constants.SITEFOOTER_*</c>). Those GUIDs are node <em>titles</em>, not ids - SharePoint
    /// uses them as well-known markers inside an otherwise ordinary node tree.</para>
    /// <para><b>Writing goes through <c>SaveMenuState</c> rather than node-by-node.</b> PnP Framework
    /// added each link as a CSOM navigation node under a parent it had fetched by id, which meant one
    /// round trip per link, a reverse-order loop to compensate for <c>Children.Add</c> prepending,
    /// and repeated <c>PendingRequestCount</c> checks. PnP Core exposes navigation nodes only for
    /// quick launch and the top bar, so that route is not available - and the endpoint the SharePoint
    /// UI itself uses takes the whole tree in one POST, which removes the ordering workaround as
    /// well.</para>
    /// <para><b>Communication sites only.</b> Other site types have no modern footer, and both
    /// halves silently no-op there, so the handler stops early rather than pretending to succeed.</para>
    /// <para><b>Not yet ported:</b> per-language footer link titles
    /// (<c>PersistMultiLanguageResources</c>) and <c>ReplaceFileUniqueToken</c> url rewriting. Both
    /// belong to handlers that land later - <c>ObjectLocalization</c> in phase 6 and
    /// <c>ObjectNavigation</c> in phase 7.</para>
    /// </remarks>
    internal class ObjectSiteFooterSettings : ObjectHandlerBase
    {
        private static readonly JsonSerializerOptions MenuStateJsonOptions = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            Converters = { new VerboseCollectionConverter() },
        };

        public override string Name => "Site Footer";

        public override string InternalName => "SiteFooter";

        public override bool WillExtract(PnPContext context, ProvisioningTemplate template, ExtractConfiguration configuration)
        {
            // The communication site check needs a round trip, which a synchronous predicate the
            // engine calls twice per run should not be doing. It happens inside the async methods.
            return true;
        }

        public override bool WillProvision(PnPContext context, ProvisioningTemplate template, ApplyConfiguration configuration)
        {
            return template.Footer != null;
        }

        #region Extract

        public override async Task<ProvisioningTemplate> ExtractObjectsAsync(PnPContext context, ProvisioningTemplate template, ExtractConfiguration configuration)
        {
            using (context.Logger?.BeginScope(Name))
            {
                if (!await SiteTypeHelper.IsCommunicationSiteAsync(context).ConfigureAwait(false))
                {
                    return template;
                }

                IWeb web = await context.Web.GetAsync(w => w.ServerRelativeUrl, w => w.Url).ConfigureAwait(false);
                await context.Site.LoadAsync(s => s.ServerRelativeUrl).ConfigureAwait(false);

                IChromeOptions chrome = await context.Web.GetBrandingManager().GetChromeOptionsAsync().ConfigureAwait(false);

                var footer = new FooterModel
                {
                    Enabled = chrome.Footer.Enabled,
                    DisplayName = chrome.Footer.DisplayName,
                    Layout = ToTemplateLayout(chrome.Footer.Layout),
                    BackgroundEmphasis = ToTemplateEmphasis(chrome.Footer.Emphasis),
                    RemoveExistingNodes = configuration?.SiteFooter?.RemoveExistingNodes ?? false,
                };

                MenuState menuState = await GetMenuStateAsync(context).ConfigureAwait(false);

                ReadTitleAndLogo(footer, menuState, web.ServerRelativeUrl);
                ReadLinks(footer, menuState, template, web.ServerRelativeUrl);

                template.Footer = footer;

                // MIGRATION PHASE 6: PnP Framework also persisted the footer logo into the template's
                // connector when PersistBrandingFiles was set, and re-read every link title in each
                // supported language. Both need the file and resource plumbing that lands with
                // ObjectFiles and ObjectLocalization.

                return template;
            }
        }

        /// <summary>
        /// Reads the footer's display name and logo out of the menu tree.
        /// </summary>
        /// <remarks>
        /// The name lives one level down, as the single child of the title marker node - not on the
        /// marker itself. The logo can come from either the title node's child (as a url) or from
        /// its own marker node, and SharePoint uses whichever the footer was configured through.
        /// </remarks>
        private void ReadTitleAndLogo(FooterModel footer, MenuState menuState, string webServerRelativeUrl)
        {
            if (menuState.Nodes.Count == 0)
            {
                return;
            }

            MenuNode titleNode = menuState.Nodes.FirstOrDefault(n => n.Title == Constants.SITEFOOTER_TITLENODEKEY);
            MenuNode titleChild = titleNode?.Nodes.FirstOrDefault();

            if (titleChild != null)
            {
                if (!string.IsNullOrEmpty(titleChild.SimpleUrl))
                {
                    footer.Logo = Tokenize(titleChild.SimpleUrl, webServerRelativeUrl);
                }

                if (!string.IsNullOrEmpty(titleChild.Title))
                {
                    footer.Name = titleChild.Title;
                }
            }

            if (string.IsNullOrEmpty(footer.Logo))
            {
                MenuNode logoNode = menuState.Nodes.FirstOrDefault(n => n.Title == Constants.SITEFOOTER_LOGONODEKEY);
                if (logoNode != null)
                {
                    footer.Logo = Tokenize(logoNode.SimpleUrl, webServerRelativeUrl);
                }
            }
        }

        private void ReadLinks(FooterModel footer, MenuState menuState, ProvisioningTemplate template, string webServerRelativeUrl)
        {
            MenuNode menuNode = menuState.Nodes.FirstOrDefault(n => n.Title == Constants.SITEFOOTER_MENUNODEKEY);
            if (menuNode == null)
            {
                return;
            }

            foreach (MenuNode node in menuNode.Nodes)
            {
                footer.FooterLinks.Add(ToFooterLink(node, template, webServerRelativeUrl));
            }
        }

        private SiteFooterLink ToFooterLink(MenuNode node, ProvisioningTemplate template, string webServerRelativeUrl)
        {
            var link = new SiteFooterLink
            {
                DisplayName = node.Title,
                Url = Tokenize(node.SimpleUrl, webServerRelativeUrl),
            };

            if (node.Nodes.Count > 0)
            {
                link.FooterLinks = new SiteFooterLinkCollection(template);
                foreach (MenuNode child in node.Nodes)
                {
                    link.FooterLinks.Add(ToFooterLink(child, template, webServerRelativeUrl));
                }
            }

            return link;
        }

        #endregion

        #region Apply

        public override async Task<TokenParser> ProvisionObjectsAsync(PnPContext context, ProvisioningTemplate template, TokenParser parser, ApplyConfiguration configuration)
        {
            using (context.Logger?.BeginScope(Name))
            {
                if (template.Footer == null)
                {
                    return parser;
                }

                if (!await SiteTypeHelper.IsCommunicationSiteAsync(context).ConfigureAwait(false))
                {
                    const string message = "Only communication sites have a modern footer - the template's footer settings were not applied.";
                    context.Logger?.LogWarning("{Source}: {Message}", Constants.LOGGING_SOURCE, message);
                    WriteMessage(message, ProvisioningMessageType.Warning);
                    return parser;
                }

                FooterModel footer = template.Footer;

                IBrandingManager branding = context.Web.GetBrandingManager();
                IChromeOptions chrome = await branding.GetChromeOptionsAsync().ConfigureAwait(false);

                chrome.Footer.Enabled = footer.Enabled;

                // A null display name on an enabled footer is rejected, and a template that simply
                // does not mention it should not be able to cause that.
                if (footer.DisplayName != null)
                {
                    chrome.Footer.DisplayName = footer.DisplayName;
                }

                chrome.Footer.Layout = ToCoreLayout(footer.Layout);
                chrome.Footer.Emphasis = ToCoreEmphasis(footer.BackgroundEmphasis);

                await branding.SetChromeOptionsAsync(chrome).ConfigureAwait(false);

                if (footer.Enabled)
                {
                    await ApplyMenuStateAsync(context, footer, parser).ConfigureAwait(false);
                }

                return parser;
            }
        }

        /// <summary>
        /// Rewrites the footer's navigation tree in a single <c>SaveMenuState</c> POST.
        /// </summary>
        private async Task ApplyMenuStateAsync(PnPContext context, FooterModel footer, TokenParser parser)
        {
            IWeb web = await context.Web.GetAsync(w => w.ServerRelativeUrl).ConfigureAwait(false);

            MenuState menuState = await GetMenuStateAsync(context).ConfigureAwait(false);

            // A site whose footer has never been configured has no tree at all. Posting an empty one
            // creates it; SharePoint then hands back a starting node key on the next read.
            if (string.IsNullOrEmpty(menuState.StartingNodeKey))
            {
                await SaveMenuStateAsync(context, new MenuState
                {
                    Version = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss:Z", CultureInfo.InvariantCulture),
                    StartingNodeTitle = Constants.SITEFOOTER_MENUNODEKEY,
                    SPSitePrefix = "/",
                    SPWebPrefix = web.ServerRelativeUrl,
                    FriendlyUrlPrefix = string.Empty,
                    SimpleUrl = string.Empty,
                }).ConfigureAwait(false);

                menuState = await GetMenuStateAsync(context).ConfigureAwait(false);
            }

            BuildMenuNode(menuState, footer, parser);
            BuildLogoNode(menuState, footer, parser);
            BuildTitleNode(menuState, footer, parser);

            menuState.Version = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss:Z", CultureInfo.InvariantCulture);
            menuState.SPWebPrefix = web.ServerRelativeUrl;

            await SaveMenuStateAsync(context, menuState).ConfigureAwait(false);
        }

        /// <summary>
        /// Puts the template's links under the menu marker node.
        /// </summary>
        /// <remarks>
        /// Links are appended in template order. PnP Framework iterated backwards because CSOM's
        /// <c>Children.Add</c> inserted at the front; writing the array directly makes that
        /// workaround unnecessary - and a reader who does not know why the loop counted down would
        /// reasonably "fix" it into a bug.
        /// </remarks>
        private static void BuildMenuNode(MenuState menuState, FooterModel footer, TokenParser parser)
        {
            MenuNode menuNode = menuState.Nodes.FirstOrDefault(n => n.Title == Constants.SITEFOOTER_MENUNODEKEY);

            if (menuNode == null)
            {
                menuNode = new MenuNode { Title = Constants.SITEFOOTER_MENUNODEKEY };
                menuState.Nodes.Add(menuNode);
            }
            else if (footer.RemoveExistingNodes)
            {
                menuNode.Nodes.Clear();
            }

            if (footer.FooterLinks == null)
            {
                return;
            }

            foreach (SiteFooterLink link in footer.FooterLinks)
            {
                menuNode.Nodes.Add(ToMenuNode(link, parser));
            }
        }

        private static MenuNode ToMenuNode(SiteFooterLink link, TokenParser parser)
        {
            var node = new MenuNode
            {
                Title = parser.ParseString(link.DisplayName),
                SimpleUrl = parser.ParseString(link.Url),
            };

            if (link.FooterLinks != null)
            {
                foreach (SiteFooterLink child in link.FooterLinks)
                {
                    node.Nodes.Add(ToMenuNode(child, parser));
                }
            }

            return node;
        }

        private static void BuildLogoNode(MenuState menuState, FooterModel footer, TokenParser parser)
        {
            MenuNode logoNode = menuState.Nodes.FirstOrDefault(n => n.Title == Constants.SITEFOOTER_LOGONODEKEY);
            string logo = parser.ParseString(footer.Logo);

            if (string.IsNullOrEmpty(logo))
            {
                // An empty logo in the template means "remove it", not "leave it alone" - that is
                // how a delta template clears one.
                if (logoNode != null)
                {
                    menuState.Nodes.Remove(logoNode);
                }

                return;
            }

            if (logoNode == null)
            {
                menuState.Nodes.Add(new MenuNode { Title = Constants.SITEFOOTER_LOGONODEKEY, SimpleUrl = logo });
            }
            else
            {
                logoNode.SimpleUrl = logo;
            }
        }

        /// <summary>
        /// Sets the footer's display name, which lives as the single child of the title marker node.
        /// </summary>
        private static void BuildTitleNode(MenuState menuState, FooterModel footer, TokenParser parser)
        {
            MenuNode titleNode = menuState.Nodes.FirstOrDefault(n => n.Title == Constants.SITEFOOTER_TITLENODEKEY);
            string name = parser.ParseString(footer.Name);

            if (string.IsNullOrEmpty(name))
            {
                if (titleNode != null)
                {
                    menuState.Nodes.Remove(titleNode);
                }

                return;
            }

            if (titleNode == null)
            {
                titleNode = new MenuNode { Title = Constants.SITEFOOTER_TITLENODEKEY };
                menuState.Nodes.Add(titleNode);
            }

            if (titleNode.Nodes.Count == 0)
            {
                titleNode.Nodes.Add(new MenuNode());
            }

            titleNode.Nodes[0].Title = name;
        }

        #endregion

        #region The MenuState endpoints

        private static async Task<MenuState> GetMenuStateAsync(PnPContext context)
        {
            ApiRequestResponse response = await context.Web.ExecuteRequestAsync(new ApiRequest(
                ApiRequestType.SPORest,
                $"_api/navigation/MenuState?menuNodeKey='{Constants.SITEFOOTER_NODEKEY}'")).ConfigureAwait(false);

            return ParseMenuState(response.Response);
        }

        /// <summary>
        /// Pulls the menu state out of a response whose shape depends on the OData mode in force.
        /// </summary>
        /// <remarks>
        /// <para>Verified against the tenant rather than assumed: the endpoint answers
        /// <c>{"d":{"MenuState":{…}}}</c> in verbose mode, and <c>{"d":{"MenuState":null}}</c> on a
        /// site whose footer has never been configured. A parser that reads the root, or that
        /// unwraps only one level, gets an empty object either way and reports it as "no footer" -
        /// the exact silent-default failure that made a working regional settings write look broken
        /// earlier in this migration.</para>
        /// <para>So the wrappers are peeled until an object that actually looks like a menu state
        /// appears, and a genuine <c>null</c> becomes an empty state with no starting node key -
        /// which is the signal the caller uses to create the tree.</para>
        /// </remarks>
        private static MenuState ParseMenuState(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return new MenuState();
            }

            using (JsonDocument document = JsonDocument.Parse(json))
            {
                JsonElement root = document.RootElement;

                for (int depth = 0; depth < 4 && root.ValueKind == JsonValueKind.Object && !IsMenuState(root); depth++)
                {
                    bool unwrapped = false;

                    foreach (string wrapper in new[] { "d", "MenuState", "menuState", "value" })
                    {
                        if (!root.TryGetProperty(wrapper, out JsonElement inner))
                        {
                            continue;
                        }

                        if (inner.ValueKind == JsonValueKind.Null)
                        {
                            // The footer has no navigation tree yet.
                            return new MenuState();
                        }

                        if (inner.ValueKind == JsonValueKind.Object)
                        {
                            root = inner;
                            unwrapped = true;
                            break;
                        }
                    }

                    if (!unwrapped)
                    {
                        break;
                    }
                }

                if (root.ValueKind != JsonValueKind.Object || !IsMenuState(root))
                {
                    return new MenuState();
                }

                return JsonSerializer.Deserialize<MenuState>(root.GetRawText(), MenuStateJsonOptions) ?? new MenuState();
            }
        }

        private static bool IsMenuState(JsonElement element)
        {
            return element.TryGetProperty("Nodes", out _) || element.TryGetProperty("StartingNodeKey", out _);
        }

        private static async Task SaveMenuStateAsync(PnPContext context, MenuState menuState)
        {
            string body = JsonSerializer.Serialize(new MenuStateWrapper { MenuState = menuState }, MenuStateJsonOptions);

            await context.Web.ExecuteRequestAsync(new ApiRequest(
                HttpMethod.Post, ApiRequestType.SPORest, "_api/navigation/SaveMenuState", body)).ConfigureAwait(false);
        }

        #endregion

        #region Enum mapping

        private static SiteFooterLayout ToTemplateLayout(FooterLayoutType layout)
        {
            return Enum.TryParse(layout.ToString(), out SiteFooterLayout parsed) ? parsed : SiteFooterLayout.Simple;
        }

        private static FooterLayoutType ToCoreLayout(SiteFooterLayout layout)
        {
            return Enum.TryParse(layout.ToString(), out FooterLayoutType parsed) ? parsed : FooterLayoutType.Simple;
        }

        /// <summary>
        /// Maps the footer emphasis by name.
        /// </summary>
        /// <remarks>
        /// <b>Never cast these.</b> <c>FooterVariantThemeType</c> numbers its members in the opposite
        /// order from the template's <c>Emphasis</c> - <c>Strong</c> is 0 there and 3 here - so a
        /// cast compiles, runs, and silently produces the wrong colour.
        /// </remarks>
        private static Emphasis ToTemplateEmphasis(FooterVariantThemeType emphasis)
        {
            return Enum.TryParse(emphasis.ToString(), out Emphasis parsed) ? parsed : Emphasis.None;
        }

        private static FooterVariantThemeType ToCoreEmphasis(Emphasis emphasis)
        {
            return Enum.TryParse(emphasis.ToString(), out FooterVariantThemeType parsed) ? parsed : FooterVariantThemeType.None;
        }

        #endregion

        #region The MenuState payload

        /// <summary>
        /// The navigation tree behind the modern footer, as <c>MenuState</c> returns it.
        /// </summary>
        private sealed class MenuState
        {
            public string FriendlyUrlPrefix { get; set; }

            public List<MenuNode> Nodes { get; set; } = new List<MenuNode>();

            public string SimpleUrl { get; set; }

            public string SPSitePrefix { get; set; }

            public string SPWebPrefix { get; set; }

            public string StartingNodeKey { get; set; }

            public string StartingNodeTitle { get; set; }

            public string Version { get; set; }
        }

        private sealed class MenuNode
        {
            public string FriendlyUrlSegment { get; set; }

            public bool IsDeleted { get; set; }

            public bool IsHidden { get; set; }

            /// <summary>
            /// SharePoint's id for an existing node. Left null on a node this run is creating -
            /// preserving it on nodes that already exist is what stops <c>SaveMenuState</c> from
            /// recreating (and so reordering) the whole tree.
            /// </summary>
            public string Key { get; set; }

            public List<MenuNode> Nodes { get; set; } = new List<MenuNode>();

            public int NodeType { get; set; }

            public string SimpleUrl { get; set; }

            public string Title { get; set; }
        }

        private sealed class MenuStateWrapper
        {
            [JsonPropertyName("menuState")]
            public MenuState MenuState { get; set; }
        }

        /// <summary>
        /// Reads and writes a node list in SharePoint's verbose OData collection form.
        /// </summary>
        /// <remarks>
        /// <para>This endpoint speaks verbose OData, where a collection is not a bare array but
        /// <c>{"results": [ … ]}</c>. Posting a bare array is rejected with <i>"A node of type
        /// 'StartArray' was read from the JSON reader when trying to read a value of a property"</i>
        /// - which names the symptom and not the cause, and is easy to misread as a malformed
        /// node.</para>
        /// <para>Reading accepts either form, because the wrapper is a function of the OData mode
        /// the request happened to negotiate and not of the data.</para>
        /// </remarks>
        private sealed class VerboseCollectionConverter : JsonConverter<List<MenuNode>>
        {
            public override List<MenuNode> Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
            {
                if (reader.TokenType == JsonTokenType.Null)
                {
                    return new List<MenuNode>();
                }

                using (JsonDocument document = JsonDocument.ParseValue(ref reader))
                {
                    JsonElement element = document.RootElement;

                    if (element.ValueKind == JsonValueKind.Object && element.TryGetProperty("results", out JsonElement results))
                    {
                        element = results;
                    }

                    if (element.ValueKind != JsonValueKind.Array)
                    {
                        return new List<MenuNode>();
                    }

                    var nodes = new List<MenuNode>();
                    foreach (JsonElement item in element.EnumerateArray())
                    {
                        // Deserializing a MenuNode re-enters this converter only through that
                        // node's own Nodes property, which is exactly what should happen.
                        MenuNode node = JsonSerializer.Deserialize<MenuNode>(item.GetRawText(), options);
                        if (node != null)
                        {
                            nodes.Add(node);
                        }
                    }

                    return nodes;
                }
            }

            public override void Write(Utf8JsonWriter writer, List<MenuNode> value, JsonSerializerOptions options)
            {
                writer.WriteStartObject();
                writer.WritePropertyName("results");
                writer.WriteStartArray();

                foreach (MenuNode node in value ?? new List<MenuNode>())
                {
                    JsonSerializer.Serialize(writer, node, options);
                }

                writer.WriteEndArray();
                writer.WriteEndObject();
            }
        }

        #endregion
    }
}
