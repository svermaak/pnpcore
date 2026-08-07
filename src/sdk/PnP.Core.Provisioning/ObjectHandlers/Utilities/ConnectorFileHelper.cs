using PnP.Core.Provisioning.Connectors;
using System;
using System.IO;

namespace PnP.Core.Provisioning.ObjectHandlers.Utilities
{
    /// <summary>
    /// Reads a file out of a template's connector by its template-relative path.
    /// </summary>
    /// <remarks>
    /// A template's <c>Src</c> values are paths relative to whatever the connector considers its
    /// container - a folder on disk, a folder in a <c>.pnp</c> package, a document library. This
    /// splits such a path into the container and file name the connector API expects.
    /// </remarks>
    internal static class ConnectorFileHelper
    {
        /// <summary>
        /// Reads a file from the connector.
        /// </summary>
        /// <param name="connector">The template's file connector</param>
        /// <param name="fileName">Path to the file, relative to the connector's container</param>
        /// <returns>The file's bytes</returns>
        /// <exception cref="ArgumentNullException"><paramref name="connector"/> is null</exception>
        /// <exception cref="ArgumentException">No such file in the connector</exception>
        internal static byte[] GetFileBytes(FileConnectorBase connector, string fileName)
        {
            if (connector == null)
            {
                throw new ArgumentNullException(nameof(connector));
            }

            string container = string.Empty;
            if (fileName.Contains("\\") || fileName.Contains("/"))
            {
                string normalized = fileName.Replace("/", "\\");
                container = fileName.Substring(0, normalized.LastIndexOf("\\", StringComparison.Ordinal));
                fileName = fileName.Substring(normalized.LastIndexOf("\\", StringComparison.Ordinal) + 1);
            }

            // Prefix the connector's own container, when it has one
            if (!string.IsNullOrEmpty(container))
            {
                if (!string.IsNullOrEmpty(connector.GetContainer()))
                {
                    container = $@"{connector.GetContainer()}\{container.TrimStart('/')}";
                }
            }
            else
            {
                container = connector.GetContainer();
            }

            Stream stream = connector.GetFileStream(fileName, container);
            if (stream == null)
            {
                // The path may have been url encoded on the way in
                fileName = Uri.UnescapeDataString(fileName);
                stream = connector.GetFileStream(fileName, container);
            }

            if (stream == null)
            {
                throw new ArgumentException($"The specified filename '{fileName}' cannot be found", nameof(fileName));
            }

            using (stream)
            {
                using (var buffer = new MemoryStream())
                {
                    stream.CopyTo(buffer);
                    return buffer.ToArray();
                }
            }
        }
    }
}
