using System;

namespace PnP.Core.Provisioning.Extensibility
{
    /// <summary>
    /// Thrown when a custom extensibility handler fails during a provisioning run.
    /// </summary>
    /// <remarks>
    /// PnP Framework marked this <c>[Serializable]</c> with a <c>SerializationInfo</c> constructor.
    /// Both are dropped here: BinaryFormatter-based exception serialization is obsolete from
    /// .NET 8 (SYSLIB0051) and disabled by default, so keeping them would add a warning in exchange
    /// for a code path nothing can use. The rest of PnP Core's exception types do the same.
    /// </remarks>
    public sealed class ExtensiblityPipelineException : Exception
    {
        /// <summary>
        /// Creates the exception with a system supplied message.
        /// </summary>
        public ExtensiblityPipelineException() : base()
        {
        }

        /// <summary>
        /// Creates the exception with the given message.
        /// </summary>
        /// <param name="message">A string that describes the exception.</param>
        public ExtensiblityPipelineException(string message) : base(message)
        {
        }

        /// <summary>
        /// Creates the exception with the given message and the exception that caused it.
        /// </summary>
        /// <param name="message">A string that describes the exception.</param>
        /// <param name="innerException">The exception that is the cause of the current exception.</param>
        public ExtensiblityPipelineException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
