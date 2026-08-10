using Microsoft.Extensions.Logging;
using PnP.Core.Provisioning.Model;
using PnP.Core.Provisioning.Model.Configuration;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace PnP.Core.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Writes the resource files behind every <c>{res:…}</c> token the other handlers collected.
    /// </summary>
    /// <remarks>
    /// <para><b>It runs last on extract, and it produces nothing of its own.</b> Every other handler
    /// that meets a localizable title records the per-language values against a token key; this one
    /// turns that accumulated list into one <c>.resx</c> per language, saves them through the
    /// template's connector, and adds the matching <c>&lt;pnp:Localization&gt;</c> entries. Register
    /// it before the handlers that fill the list and the files come out empty.</para>
    /// <para><b>Apply does nothing here.</b> The resource files are read by the token parser when it
    /// is built, long before any handler runs - see <c>TokenParser.AddResourceTokens</c> - and the
    /// values are written by whichever handler owns the object. There is no separate localization
    /// pass on the way in.</para>
    /// </remarks>
    internal class ObjectLocalization : ObjectHandlerBase
    {
        public override string Name => "Localization";

        public override string InternalName => "Localization";

        public override bool WillProvision(PnPContext context, ProvisioningTemplate template, ApplyConfiguration configuration)
        {
            return false;
        }

        public override bool WillExtract(PnPContext context, ProvisioningTemplate template, ExtractConfiguration configuration)
        {
            _willExtract ??= configuration?.ToCreationInformation()?.PersistMultiLanguageResources ?? false;
            return _willExtract.Value;
        }

        public override Task<TokenParser> ProvisionObjectsAsync(PnPContext context, ProvisioningTemplate template, TokenParser parser, ApplyConfiguration configuration)
        {
            // See the remarks on the class - there is no apply pass.
            return Task.FromResult(parser);
        }

        public override Task<ProvisioningTemplate> ExtractObjectsAsync(PnPContext context, ProvisioningTemplate template, ExtractConfiguration configuration)
        {
            using (context.Logger?.BeginScope(Name))
            {
                ProvisioningTemplateCreationInformation creationInformation = configuration?.ToCreationInformation();

                if (creationInformation?.PersistMultiLanguageResources != true)
                {
                    return Task.FromResult(template);
                }

                List<Tuple<string, int, string>> resourceTokens = creationInformation.ResourceTokens;

                if (resourceTokens.Count == 0)
                {
                    // Said out loud, because the usual cause is an ordering or configuration mistake
                    // rather than a site with nothing to localize: the handlers that collect values
                    // iterate template.SupportedUILanguages, so extracting without that handler
                    // produces an empty resource set and no error anywhere.
                    string message = template.SupportedUILanguages.Count == 0
                        ? "Multi-language resources were requested but the template lists no supported UI languages, so nothing was collected. " +
                          "Include the SupportedUILanguages handler in the extract - the localized values are read per language."
                        : "Multi-language resources were requested but no localized values were found on this site.";

                    context.Logger?.LogWarning("{Source}: {Message}", Constants.LOGGING_SOURCE, message);
                    WriteMessage(message, ProvisioningMessageType.Warning);

                    return Task.FromResult(template);
                }

                if (creationInformation.FileConnector == null)
                {
                    const string message = "Multi-language resources were requested but the extract has no connector, so the resource files could not be saved.";
                    context.Logger?.LogError("{Source}: {Message}", Constants.LOGGING_SOURCE, message);
                    WriteMessage(message, ProvisioningMessageType.Error);
                    return Task.FromResult(template);
                }

                string prefix = string.IsNullOrEmpty(creationInformation.ResourceFilePrefix)
                    ? "PnPResources"
                    : creationInformation.ResourceFilePrefix;

                foreach (int lcid in resourceTokens.Select(t => t.Item2).Distinct())
                {
                    WriteResourceFile(context, template, creationInformation, resourceTokens, lcid, prefix);
                }

                return Task.FromResult(template);
            }
        }

        /// <summary>
        /// Writes one language's resource file and registers it on the template.
        /// </summary>
        private void WriteResourceFile(PnPContext context, ProvisioningTemplate template,
            ProvisioningTemplateCreationInformation creationInformation, List<Tuple<string, int, string>> resourceTokens,
            int lcid, string prefix)
        {
            var culture = new CultureInfo(lcid);
            string fileName = $"{prefix}.{culture.Name}.resx";

            // Later wins, so a handler that revisits a key overwrites rather than duplicating it -
            // a resx with two entries of the same name will not load.
            var entries = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (Tuple<string, int, string> token in resourceTokens.Where(t => t.Item2 == lcid))
            {
                entries[token.Item1] = token.Item3;
            }

            // MIGRATION NOTE: PnP Framework merged in any file of the same name it found in the
            // machine's temp folder before writing. That was reading a leftover from an earlier run
            // in the same process - and through ResourceReader, which reads binary .resources and
            // would have thrown on a .resx. It is dropped rather than reproduced: the token list
            // already accumulates across the whole extract, which is what the merge was reaching for.

            byte[] content = Encoding.UTF8.GetBytes(BuildResx(entries));

            using (var stream = new MemoryStream(content))
            {
                creationInformation.FileConnector.SaveFileStream(fileName, stream);
            }

            template.Localizations.Add(new Localization
            {
                LCID = lcid,
                Name = culture.NativeName,
                ResourceFile = fileName,
            });

            context.Logger?.LogInformation("{Source}: wrote {Count} value(s) to {File}.",
                Constants.LOGGING_SOURCE, entries.Count, fileName);
        }

        /// <summary>
        /// Builds a resource file.
        /// </summary>
        /// <remarks>
        /// <para>Emitted as plain XML rather than through <c>ResXResourceWriter</c>, which lives in a
        /// package PnP Core does not reference and which PnP Framework only pulled in because the
        /// framework's own <c>ResourceWriter</c> produced <b>binary</b> <c>.resources</c> files that
        /// no editor could open.</para>
        /// <para>The header block is boilerplate that every resx carries, and it is what makes the
        /// file openable in Visual Studio as well as by this engine - the reader here only needs the
        /// <c>data</c> elements.</para>
        /// </remarks>
        private static string BuildResx(Dictionary<string, string> entries)
        {
            XNamespace xml = "http://www.w3.org/XML/1998/namespace";

            var root = new XElement("root",
                new XElement("resheader", new XAttribute("name", "resmimetype"), new XElement("value", "text/microsoft-resx")),
                new XElement("resheader", new XAttribute("name", "version"), new XElement("value", "2.0")),
                new XElement("resheader", new XAttribute("name", "reader"),
                    new XElement("value", "System.Resources.ResXResourceReader, System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a2c561934e089")),
                new XElement("resheader", new XAttribute("name", "writer"),
                    new XElement("value", "System.Resources.ResXResourceWriter, System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a2c561934e089")));

            foreach (KeyValuePair<string, string> entry in entries.OrderBy(e => e.Key, StringComparer.Ordinal))
            {
                root.Add(new XElement("data",
                    new XAttribute("name", entry.Key),

                    // Titles legitimately have leading or trailing spaces, and a resx loses them
                    // without this - which is why real resx files carry it on every data element.
                    new XAttribute(xml + "space", "preserve"),
                    new XElement("value", entry.Value ?? string.Empty)));
            }

            return new XDocument(new XDeclaration("1.0", "utf-8", null), root).ToString();
        }
    }
}
