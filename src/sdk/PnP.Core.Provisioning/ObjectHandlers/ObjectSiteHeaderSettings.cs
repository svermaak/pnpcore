using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.Services;
using System;
using System.Threading.Tasks;
using HeaderModel = PnP.Core.Provisioning.Model.SiteHeader;

namespace PnP.Core.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Reads and writes the modern site header: layout, emphasis, menu style, and whether the
    /// title and navigation are shown.
    /// </summary>
    /// <remarks>
    /// <para><b>This handler got simpler, not harder.</b> PnP Framework needed three mechanisms for
    /// five settings - CSOM properties for the layout and emphasis, a hand-rolled POST to
    /// <c>SetChromeOptions</c> because those properties do not commit on their own, and a bridge
    /// into PnP Core for the title and navigation flags. <see cref="IBrandingManager"/> covers all
    /// five, so the port is one read and one write.</para>
    /// <para>The two enum mappings are done <b>by name</b>. The template's <c>Emphasis</c> and PnP
    /// Core's <c>VariantThemeType</c> happen to agree numerically today, but the footer's
    /// equivalent does not (<c>FooterVariantThemeType.Strong</c> is 0, not 3) - so casting would be
    /// a trap set for whoever copies this file next.</para>
    /// </remarks>
    internal class ObjectSiteHeaderSettings : ObjectHandlerBase
    {
        public override string Name => "Site Header";

        public override string InternalName => "SiteHeader";

        public override bool WillExtract(PnPContext context, ProvisioningTemplate template, ExtractConfiguration configuration)
        {
            return true;
        }

        public override bool WillProvision(PnPContext context, ProvisioningTemplate template, ApplyConfiguration configuration)
        {
            return template.Header != null;
        }

        public override async Task<ProvisioningTemplate> ExtractObjectsAsync(PnPContext context, ProvisioningTemplate template, ExtractConfiguration configuration)
        {
            using (context.Logger?.BeginScope(Name))
            {
                IChromeOptions chrome = await context.Web.GetBrandingManager().GetChromeOptionsAsync().ConfigureAwait(false);

                template.Header = new HeaderModel
                {
                    Layout = ToTemplateLayout(chrome.Header.Layout),
                    BackgroundEmphasis = ToTemplateEmphasis(chrome.Header.Emphasis),
                    MenuStyle = chrome.Navigation.MegaMenuEnabled ? SiteHeaderMenuStyle.MegaMenu : SiteHeaderMenuStyle.Cascading,
                    ShowSiteTitle = !chrome.Header.HideTitle,
                    ShowSiteNavigation = chrome.Navigation.Visible,
                };

                return template;
            }
        }

        public override async Task<TokenParser> ProvisionObjectsAsync(PnPContext context, ProvisioningTemplate template, TokenParser parser, ApplyConfiguration configuration)
        {
            using (context.Logger?.BeginScope(Name))
            {
                if (template.Header == null)
                {
                    return parser;
                }

                IBrandingManager branding = context.Web.GetBrandingManager();
                IChromeOptions chrome = await branding.GetChromeOptionsAsync().ConfigureAwait(false);

                chrome.Header.Layout = ToCoreLayout(template.Header.Layout);
                chrome.Header.Emphasis = ToCoreEmphasis(template.Header.BackgroundEmphasis);
                chrome.Header.HideTitle = !template.Header.ShowSiteTitle;
                chrome.Navigation.MegaMenuEnabled = template.Header.MenuStyle == SiteHeaderMenuStyle.MegaMenu;
                chrome.Navigation.Visible = template.Header.ShowSiteNavigation;

                await branding.SetChromeOptionsAsync(chrome).ConfigureAwait(false);

                return parser;
            }
        }

        /// <summary>
        /// Maps PnP Core's header layout onto the template's.
        /// </summary>
        /// <remarks>
        /// <c>HeaderLayoutType.None</c> has no counterpart in the schema. It is the deprecated
        /// original header, which SharePoint itself renders as standard, so that is what it becomes.
        /// </remarks>
        private static SiteHeaderLayout ToTemplateLayout(HeaderLayoutType layout)
        {
            switch (layout)
            {
                case HeaderLayoutType.Compact:
                    return SiteHeaderLayout.Compact;
                case HeaderLayoutType.Minimal:
                    return SiteHeaderLayout.Minimal;
                case HeaderLayoutType.Extended:
                    return SiteHeaderLayout.Extended;
                default:
                    return SiteHeaderLayout.Standard;
            }
        }

        private static HeaderLayoutType ToCoreLayout(SiteHeaderLayout layout)
        {
            switch (layout)
            {
                case SiteHeaderLayout.Compact:
                    return HeaderLayoutType.Compact;
                case SiteHeaderLayout.Minimal:
                    return HeaderLayoutType.Minimal;
                case SiteHeaderLayout.Extended:
                    return HeaderLayoutType.Extended;
                default:
                    return HeaderLayoutType.Standard;
            }
        }

        private static Emphasis ToTemplateEmphasis(VariantThemeType emphasis)
        {
            return Enum.TryParse(emphasis.ToString(), out Emphasis parsed) ? parsed : Emphasis.None;
        }

        private static VariantThemeType ToCoreEmphasis(Emphasis emphasis)
        {
            return Enum.TryParse(emphasis.ToString(), out VariantThemeType parsed) ? parsed : VariantThemeType.None;
        }
    }
}
