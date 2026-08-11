using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.Provisioning.Connectors;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DataRowModel = PnP.Core.Provisioning.Model.DataRow;
using DirectoryModel = PnP.Core.Provisioning.Model.Directory;
using FieldModel = PnP.Core.Provisioning.Model.Field;
using FileLevelModel = PnP.Core.Provisioning.Model.FileLevel;
using FileModel = PnP.Core.Provisioning.Model.File;
using IODirectory = System.IO.Directory;
using IOFile = System.IO.File;

namespace PnP.Core.Provisioning.Test.Scenarios
{
    /// <summary>
    /// Scenario 4 - files, a folder tree, and list data rows.
    /// </summary>
    /// <remarks>
    /// <para>The connector-driven half of the engine. Everything here comes off disk through a
    /// <see cref="FileSystemConnector"/>, which is the path a real caller's template takes and the
    /// one the other scenarios never touch.</para>
    /// <para><b>Two files, not one, and one of them in a subfolder.</b> Uploading a single file to a
    /// library root would not distinguish "files upload" from "the folder path is walked" - and the
    /// folder walk is where <c>ObjectFiles</c> was silently failing until phase 6 caught it:
    /// <c>EnsureFolderAsync</c> from the web root asks SharePoint to create a folder <em>under the
    /// web</em>, so every upload into a library quietly did nothing.</para>
    /// <para><b>The data rows are applied twice on purpose.</b> Idempotence is the behaviour worth
    /// pinning, and it is asserted on the row <em>count</em> - an assertion that only checks the row
    /// is present stays true through a duplicate.</para>
    /// </remarks>
    [TestClass]
    public class Scenario04FilesAndDataRowsTests : ScenarioTestBase
    {
        private const string Prefix = "PnPCoreScenario4_";

        private static readonly string LibraryTitle = $"{Prefix}Docs";
        private static readonly string LibraryUrl = $"{Prefix}Docs";
        private static readonly string ListTitle = $"{Prefix}Rows";
        private static readonly string ListUrl = $"Lists/{Prefix}Rows";
        private static readonly string KeyFieldName = $"{Prefix}Key";
        private const string KeyFieldId = "{2c5e8a41-7b3d-4e9f-a1c8-6d4b2f9e3a04}";
        private const string SubFolder = "Reference";

        [TestMethod]
        [TestCategory("Live")]
        [TestCategory("Scenario")]
        [Timeout(45 * 60 * 1000)]
        public async Task Scenario4_FilesFoldersAndDataRows()
        {
            string sourceFolder = CreateSourceFiles();

            try
            {
                await RunScenarioAsync("s4", BuildTemplate(sourceFolder), new[]
                {
                    ConfigurationHandler.Lists,
                    ConfigurationHandler.Fields,
                },
                AssertAsync).ConfigureAwait(false);
            }
            finally
            {
                TryDelete(sourceFolder);
            }
        }

        /// <summary>
        /// Writes the files the template uploads, and returns the folder holding them.
        /// </summary>
        /// <remarks>
        /// Built on disk rather than committed as test assets: the content is three lines long and
        /// generating it keeps the fixture and the assertion in the same file, where a reader can see
        /// that they agree.
        /// </remarks>
        private static string CreateSourceFiles()
        {
            string folder = Path.Combine(Path.GetTempPath(), $"{Prefix}{Guid.NewGuid():N}");

            IODirectory.CreateDirectory(folder);
            IODirectory.CreateDirectory(Path.Combine(folder, SubFolder));

            IOFile.WriteAllText(Path.Combine(folder, "root.txt"), "Scenario 4 root file");
            IOFile.WriteAllText(Path.Combine(folder, SubFolder, "nested.txt"), "Scenario 4 nested file");

            return folder;
        }

        private static ProvisioningTemplate BuildTemplate(string sourceFolder)
        {
            var template = new ProvisioningTemplate
            {
                Id = "SCENARIO-4",
                Connector = new FileSystemConnector(sourceFolder, string.Empty),
            };

            template.Lists.Add(new ListInstance
            {
                Title = LibraryTitle,
                Url = LibraryUrl,
                TemplateType = (int)ListTemplateType.DocumentLibrary,

                // Publishing a file only means anything on a library with minor versions on. Without
                // these the handler correctly warns and moves on - which is right behaviour, and
                // would leave the publish path unexercised while the scenario looked green.
                EnableVersioning = true,
                EnableMinorVersions = true,
            });

            // A key column, so the data rows can be matched on re-apply rather than duplicated.
            template.SiteFields.Add(new FieldModel
            {
                SchemaXml = $"<Field ID=\"{KeyFieldId}\" Type=\"Text\" Name=\"{KeyFieldName}\" " +
                    $"StaticName=\"{KeyFieldName}\" DisplayName=\"Key\" Group=\"{Prefix}Group\" />",
            });

            var list = new ListInstance
            {
                Title = ListTitle,
                Url = ListUrl,
                TemplateType = (int)ListTemplateType.GenericList,
            };

            list.FieldRefs.Add(new FieldRef(KeyFieldName) { Id = Guid.Parse(KeyFieldId) });

            list.DataRows.KeyColumn = KeyFieldName;
            list.DataRows.UpdateBehavior = UpdateBehavior.Overwrite;

            list.DataRows.Add(Row("alpha", "Alpha original"));
            list.DataRows.Add(Row("beta", "Beta original"));

            template.Lists.Add(list);

            // Uploaded to the library root...
            template.Files.Add(new FileModel
            {
                Src = "root.txt",
                Folder = LibraryUrl,
                Overwrite = true,
                Level = FileLevelModel.Published,
            });

            // ...and into a subfolder of it, which is the case that used to fail silently.
            template.Files.Add(new FileModel
            {
                Src = $"{SubFolder}\\nested.txt",
                Folder = $"{LibraryUrl}/{SubFolder}",
                Overwrite = true,
                Level = FileLevelModel.Published,
            });

            return template;
        }

        private static DataRowModel Row(string key, string title)
        {
            var row = new DataRowModel();

            row.Values.Add("Title", title);
            row.Values.Add(KeyFieldName, key);

            return row;
        }

        private static async Task AssertAsync(ProvisioningTemplate extracted, PnPContext site)
        {
            // ------------------------------------------------------------------------------------
            // Both files uploaded, including the one in a subfolder.
            // ------------------------------------------------------------------------------------
            await site.Web.LoadAsync(w => w.ServerRelativeUrl).ConfigureAwait(false);

            string libraryRoot = $"{site.Web.ServerRelativeUrl.TrimEnd('/')}/{LibraryUrl}";

            Assert.IsTrue(await FileExistsAsync(site, $"{libraryRoot}/root.txt").ConfigureAwait(false),
                "The file at the library root was not uploaded.");

            // The one that matters: a file in a subfolder needs the folder path walked, and the
            // handler used to report success while uploading nothing.
            Assert.IsTrue(await FileExistsAsync(site, $"{libraryRoot}/{SubFolder}/nested.txt").ConfigureAwait(false),
                $"The file in '{SubFolder}' was not uploaded. This is the folder walk failing - and it " +
                "fails silently, because SharePoint accepts the request and creates nothing.");

            // ------------------------------------------------------------------------------------
            // The data rows landed, once each.
            // ------------------------------------------------------------------------------------
            await site.Web.LoadAsync(w => w.Lists.QueryProperties(l => l.Title, l => l.ItemCount))
                .ConfigureAwait(false);

            IList rows = site.Web.Lists.AsRequested().FirstOrDefault(l => l.Title == ListTitle);

            Assert.IsNotNull(rows, "The data row list was not created.");

            Console.WriteLine($"'{ListTitle}' holds {rows.ItemCount} item(s)");

            Assert.AreEqual(2, rows.ItemCount, "The template's two data rows did not land exactly once each.");

            // ------------------------------------------------------------------------------------
            // Re-applying does not duplicate them, and does update them.
            // ------------------------------------------------------------------------------------
            ProvisioningTemplate second = BuildTemplate(CreateSourceFiles());
            second.Lists.Single(l => l.Title == ListTitle).DataRows.Clear();
            second.Lists.Single(l => l.Title == ListTitle).DataRows.Add(Row("alpha", "Alpha updated"));

            await site.GetProvisioningManager().ApplyTemplateAsync(second, new ApplyConfiguration
            {
                MessagesDelegate = (message, type) => Console.WriteLine($"[{type}] {message}"),
            }).ConfigureAwait(false);

            await site.Web.LoadAsync(w => w.Lists.QueryProperties(l => l.Title, l => l.ItemCount))
                .ConfigureAwait(false);

            IList afterSecond = site.Web.Lists.AsRequested().FirstOrDefault(l => l.Title == ListTitle);

            Console.WriteLine($"After re-apply, '{ListTitle}' holds {afterSecond.ItemCount} item(s)");

            // The count is the assertion. "The row is there" stays true through a duplicate.
            Assert.AreEqual(2, afterSecond.ItemCount,
                "Re-applying the template duplicated a data row - the key column did not match it.");

            await afterSecond.LoadItemsByCamlQueryAsync("<View><Query></Query></View>").ConfigureAwait(false);

            IListItem alpha = afterSecond.Items.AsRequested()
                .FirstOrDefault(i => i.Values.TryGetValue(KeyFieldName, out object k) && (string)k == "alpha");

            Assert.IsNotNull(alpha, "The keyed row is gone after the second apply.");
            Assert.AreEqual("Alpha updated", alpha.Values["Title"]?.ToString(),
                "The matched row was not updated, so the key column matched but the update did not happen.");

            // ------------------------------------------------------------------------------------
            // And the extract found the lists.
            // ------------------------------------------------------------------------------------
            Assert.IsTrue(extracted.Lists.Any(l => l.Title == LibraryTitle),
                "The extract did not report the document library.");
            Assert.IsTrue(extracted.Lists.Any(l => l.Title == ListTitle),
                "The extract did not report the data row list.");
        }

        private static async Task<bool> FileExistsAsync(PnPContext site, string serverRelativeUrl)
        {
            try
            {
                IFile file = await site.Web.GetFileByServerRelativeUrlAsync(serverRelativeUrl)
                    .ConfigureAwait(false);

                Console.WriteLine($"Found {serverRelativeUrl}");

                return file != null;
            }
            catch (Exception)
            {
                Console.WriteLine($"Missing {serverRelativeUrl}");
                return false;
            }
        }

        private static void TryDelete(string folder)
        {
            try
            {
                IODirectory.Delete(folder, recursive: true);
            }
            catch (Exception)
            {
                // A leaked temp folder is not worth failing a scenario over.
            }
        }
    }
}
