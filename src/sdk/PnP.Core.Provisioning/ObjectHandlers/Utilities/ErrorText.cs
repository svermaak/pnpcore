using PnP.Core.Services;
using System;
using System.Collections.Generic;

namespace PnP.Core.Provisioning.ObjectHandlers.Utilities
{
    /// <summary>
    /// Renders an exception into text worth putting in a provisioning warning.
    /// </summary>
    /// <remarks>
    /// <para><c>ex.Message</c> on a PnP Core service failure is a fixed banner - "SharePoint Rest
    /// service exception" - and says nothing at all about what went wrong. What SharePoint actually
    /// said is on the <see cref="ServiceError"/>, and for a CSOM failure it is in that error's
    /// <c>ToString</c> rather than its <c>Message</c>.</para>
    /// <para>A handler that swallows an error and reports only the banner has turned a diagnosable
    /// failure into a mystery. This is the difference between "the view could not be created:
    /// SharePoint Rest service exception" and a message naming the offending attribute.</para>
    /// </remarks>
    internal static class ErrorText
    {
        /// <summary>
        /// A one-line-ish description of what failed and why.
        /// </summary>
        internal static string Describe(Exception ex)
        {
            if (ex == null)
            {
                return string.Empty;
            }

            var parts = new List<string>();

            if (ex is ServiceException serviceException && serviceException.Error is ServiceError error)
            {
                string rendered = error.ToString();
                if (!string.IsNullOrEmpty(rendered))
                {
                    parts.Add(rendered);
                }
                else if (!string.IsNullOrEmpty(error.Message))
                {
                    parts.Add(error.Message);
                }
            }

            if (parts.Count == 0)
            {
                parts.Add(ex.Message);
            }

            if (ex.InnerException != null && !string.IsNullOrEmpty(ex.InnerException.Message))
            {
                parts.Add(ex.InnerException.Message);
            }

            return string.Join(" | ", parts);
        }
    }
}
