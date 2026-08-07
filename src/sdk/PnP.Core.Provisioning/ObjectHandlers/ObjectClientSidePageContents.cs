using Microsoft.Extensions.Logging;
using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.Provisioning.ObjectHandlers.Utilities;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PnP.Core.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Extracts the site's modern pages into the template.
    /// </summary>
    /// <remarks>
    /// <para><b>Extract only.</b> This handler is the read half of the pair completed by
    /// <see cref="ObjectClientSidePages"/>, which does all the writing. It surveys the pages
    /// library, decides which pages to export, and hands each one to
    /// <see cref="ClientSidePageContentsHelper"/>.</para>
    /// <para><b>By default only the home page is exported.</b> Everything else needs
    /// <c>IncludeAllClientSidePages</c>. A site's pages are usually content rather than structure,
    /// and a template that drags every page along is rarely what the author wanted.</para>
    /// <para><b>Translations are exported after their sources</b>, because a translated page has to
    /// attach to a page that is already in the template. That is what the
    /// <c>OrderBy(IsTranslation)</c> is for - it is not cosmetic.</para>
    /// </remarks>
    internal class ObjectClientSidePageContents : ObjectContentHandlerBase
    {
        /// <summary>
        /// Every .aspx in the library, folders included, with the fields the survey needs.
        /// </summary>
        /// <remarks>
        /// <para><b>The <c>ViewFields</c> are not optional.</b> Without them SharePoint returns the
        /// library's default view fields, which do not include <c>FileRef</c>,
        /// <c>ClientSideApplicationId</c> or <c>UniqueId</c> - so every page fails the "is this a
        /// modern page" test and extraction silently reports an empty site. CSOM's
        /// <c>Context.Load(items)</c> pulled everything and hid this.</para>
        /// <para>The translation fields are added separately, because they only exist once
        /// multilingual pages are enabled - see <see cref="BuildQuery"/>.</para>
        /// </remarks>
        private const string AspxFilesQueryFormat = @"
            <View Scope='Recursive'>
              <Query>
                <Where>
                  <Contains>
                    <FieldRef Name='File_x0020_Type'/>
                    <Value Type='text'>aspx</Value>
                  </Contains>
                </Where>
              </Query>
              <ViewFields>
                <FieldRef Name='FileRef'/>
                <FieldRef Name='FileLeafRef'/>
                <FieldRef Name='ClientSideApplicationId'/>
                <FieldRef Name='UniqueId'/>
                {0}
              </ViewFields>
            </View>";

        /// <summary>
        /// The multilingual fields, which only exist on a library with translations enabled.
        /// </summary>
        private const string TranslationViewFields = @"
                <FieldRef Name='_SPIsTranslation'/>
                <FieldRef Name='_SPTranslatedLanguages'/>
                <FieldRef Name='_SPTranslationSourceItemId'/>
                <FieldRef Name='_SPTranslationLanguage'/>";

        private const string FileRefField = "FileRef";
        private const string FileLeafRefField = "FileLeafRef";
        private const string ClientSideApplicationIdField = "ClientSideApplicationId";
        private const string IsTranslationField = "_SPIsTranslation";
        private const string TranslatedLanguagesField = "_SPTranslatedLanguages";
        private const string PageIdField = "UniqueId";
        private const string TranslationSourceItemIdField = "_SPTranslationSourceItemId";
        private const string TranslationLanguageField = "_SPTranslationLanguage";

        /// <summary>
        /// The feature id every modern page carries in <c>ClientSideApplicationId</c>. A page in the
        /// library without it is a classic page and belongs to <c>ObjectPages</c> instead.
        /// </summary>
        private static readonly Guid ModernPageFeatureId = new Guid("B6917CB1-93A0-4B97-A84D-7CF49975D4EC");

        /// <summary>
        /// The entity type of the real pages library. A site with Viva Connections has a second
        /// <c>WebPageLibrary</c> ("Announcements") that must not be mistaken for it.
        /// </summary>
        private const string SitePagesEntityType = "SP.Data.SitePagesItem";

        public override string Name => "Client Side Page Contents";

        public override string InternalName => "ClientSidePageContents";

        public override bool WillProvision(PnPContext context, ProvisioningTemplate template, ApplyConfiguration configuration)
        {
            return false;
        }

        public override bool WillExtract(PnPContext context, ProvisioningTemplate template, ExtractConfiguration configuration)
        {
            return true;
        }

        public override Task<TokenParser> ProvisionObjectsAsync(PnPContext context, ProvisioningTemplate template, TokenParser parser, ApplyConfiguration configuration)
        {
            // Writing is ObjectClientSidePages' job - see the remarks on the class.
            return Task.FromResult(parser);
        }

        public override async Task<ProvisioningTemplate> ExtractObjectsAsync(PnPContext context, ProvisioningTemplate template, ExtractConfiguration configuration)
        {
            using (context.Logger?.BeginScope(Name))
            {
                IWeb web = await context.Web.GetAsync(w => w.WelcomePage, w => w.ServerRelativeUrl, w => w.Url,
                    w => w.Language, w => w.IsMultilingual).ConfigureAwait(false);

                IList sitePagesLibrary = await FindSitePagesLibraryAsync(context).ConfigureAwait(false);
                if (sitePagesLibrary == null)
                {
                    context.Logger?.LogInformation("{Source}: this site has no pages library - nothing to extract.", Constants.LOGGING_SOURCE);
                    return template;
                }

                await sitePagesLibrary.LoadAsync(l => l.RootFolder.QueryProperties(f => f.ServerRelativeUrl)).ConfigureAwait(false);
                string baseUrl = $"{sitePagesLibrary.RootFolder.ServerRelativeUrl}/";

                string templatesFolderName = await GetTemplatesFolderNameAsync(context).ConfigureAwait(false);

                await sitePagesLibrary.LoadItemsByCamlQueryAsync(BuildQuery(web.IsMultilingual)).ConfigureAwait(false);
                List<IListItem> pages = sitePagesLibrary.Items.AsRequested().ToList();

                if (pages.Count == 0)
                {
                    return template;
                }

                List<PageToExport> pagesToExport = SurveyPages(pages, web, baseUrl, templatesFolderName);
                LinkTranslationsToTheirSources(pagesToExport);

                bool includeAllPages = configuration?.Pages?.IncludeAllClientSidePages ?? false;
                var helper = new ClientSidePageContentsHelper();

                int currentPageIndex = 1;
                foreach (PageToExport page in pagesToExport.OrderBy(p => p.IsTranslation))
                {
                    if (includeAllPages || page.IsHomePage)
                    {
                        WriteSubProgress("ClientSidePage",
                            !string.IsNullOrWhiteSpace(page.PageName) ? page.PageName : page.PageUrl,
                            currentPageIndex, pagesToExport.Count);

                        await helper.ExtractClientSidePageAsync(context, template, configuration, page, MessagesDelegate).ConfigureAwait(false);
                    }

                    currentPageIndex++;
                }

                return template;
            }
        }

        /// <summary>
        /// Builds the survey query, asking for the translation fields only when they exist.
        /// </summary>
        /// <remarks>
        /// A <c>ViewFields</c> entry naming a field the library does not have makes SharePoint
        /// reject the whole query, and the multilingual fields appear only once translations are
        /// enabled - so they are requested conditionally rather than optimistically.
        /// </remarks>
        private static string BuildQuery(bool isMultilingual)
        {
            return string.Format(CultureInfo.InvariantCulture, AspxFilesQueryFormat,
                isMultilingual ? TranslationViewFields : string.Empty);
        }

        /// <summary>
        /// Finds the site's real pages library.
        /// </summary>
        private static async Task<IList> FindSitePagesLibraryAsync(PnPContext context)
        {
            try
            {
                await context.Web.LoadAsync(w => w.Lists.QueryProperties(
                    l => l.Id, l => l.Title, l => l.TemplateType, l => l.ListItemEntityTypeFullName)).ConfigureAwait(false);

                return context.Web.Lists.AsRequested()
                    .Where(l => l.TemplateType == ListTemplateType.WebPageLibrary)
                    .FirstOrDefault(l => SitePagesEntityType.Equals(l.ListItemEntityTypeFullName, StringComparison.Ordinal));
            }
            catch (Exception ex)
            {
                // An incorrectly provisioned site can make enumerating lists fail outright. Falling
                // back to the well-known title is what PnP Framework did, and it is better than
                // reporting a site with pages as having none.
                context.Logger?.LogWarning(ex, "{Source}: could not enumerate lists; falling back to the 'Site Pages' title.",
                    Constants.LOGGING_SOURCE);

                try
                {
                    return await context.Web.Lists.GetByTitleAsync("Site Pages").ConfigureAwait(false);
                }
                catch (Exception)
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// The name of the pages library's templates folder.
        /// </summary>
        /// <remarks>
        /// It is renameable, and it is only discoverable through the library's property bag - so PnP
        /// Core resolves it from a throwaway page object rather than exposing it on the list.
        /// </remarks>
        private static async Task<string> GetTemplatesFolderNameAsync(PnPContext context)
        {
            try
            {
                IPage dummyPage = await context.Web.NewPageAsync().ConfigureAwait(false);
                string folder = await dummyPage.GetTemplatesFolderAsync().ConfigureAwait(false);

                return !string.IsNullOrEmpty(folder) ? folder : "Templates";
            }
            catch (Exception)
            {
                return "Templates";
            }
        }

        /// <summary>
        /// Turns the library's list items into the export list, working out for each page whether it
        /// is the home page, a template, or a translation.
        /// </summary>
        private static List<PageToExport> SurveyPages(List<IListItem> pages, IWeb web, string baseUrl, string templatesFolderName)
        {
            var pagesToExport = new List<PageToExport>();
            string homePageUrl = web.WelcomePage;

            foreach (IListItem page in pages)
            {
                string pageUrl = GetString(page, FileRefField);
                if (string.IsNullOrEmpty(pageUrl))
                {
                    continue;
                }

                // Only modern pages belong here; a classic .aspx in the same library is ObjectPages'.
                string applicationId = GetString(page, ClientSideApplicationIdField);
                if (!ModernPageFeatureId.ToString().Equals(applicationId, StringComparison.InvariantCultureIgnoreCase))
                {
                    continue;
                }

                var pageToExport = new PageToExport
                {
                    PageUrl = pageUrl,

                    // The name includes the folder, so a page in a sub folder round trips there.
                    PageName = Regex.Replace(pageUrl, Regex.Escape(baseUrl), "", RegexOptions.IgnoreCase),
                    IsTemplate = pageUrl.IndexOf($"/{templatesFolderName}/", StringComparison.InvariantCultureIgnoreCase) > -1,
                    IsHomePage = !string.IsNullOrEmpty(homePageUrl)
                        && pageUrl.EndsWith(homePageUrl, StringComparison.InvariantCultureIgnoreCase),
                };

                // These fields only exist once multilingual pages are enabled on the library.
                if (bool.TryParse(GetString(page, IsTranslationField), out bool isTranslation))
                {
                    pageToExport.IsTranslation = isTranslation;
                }

                if (Guid.TryParse(GetString(page, PageIdField), out Guid pageId))
                {
                    pageToExport.PageId = pageId;
                }

                if (Guid.TryParse(GetString(page, TranslationSourceItemIdField), out Guid sourcePageId))
                {
                    pageToExport.SourcePageId = sourcePageId;
                }

                pageToExport.Language = GetString(page, TranslationLanguageField);

                if (page.Values.TryGetValue(TranslatedLanguagesField, out object translatedLanguages)
                    && translatedLanguages is IEnumerable<string> languages)
                {
                    pageToExport.TranslatedLanguages = languages.ToList();
                }

                pagesToExport.Add(pageToExport);
            }

            return pagesToExport;
        }

        /// <summary>
        /// Resolves each translation's source page id to the source page's name.
        /// </summary>
        /// <remarks>
        /// The template nests a translation under its source by <em>name</em>, but SharePoint links
        /// them by list item id - so the pairing has to happen while the whole set is in hand.
        /// </remarks>
        private static void LinkTranslationsToTheirSources(List<PageToExport> pagesToExport)
        {
            foreach (PageToExport page in pagesToExport.Where(p => p.IsTranslation))
            {
                page.SourcePageName = pagesToExport.FirstOrDefault(p => p.PageId == page.SourcePageId)?.PageName;
            }
        }

        private static string GetString(IListItem item, string fieldName)
        {
            return item.Values.TryGetValue(fieldName, out object value) ? value?.ToString() : null;
        }
    }
}
