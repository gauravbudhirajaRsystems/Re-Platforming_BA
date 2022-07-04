// © Copyright 2018 Levit & James, Inc.

using System;
using System.Runtime.Serialization;

namespace LevitJames.TextServices
{
    /// <summary>
    ///     An exception class used by the RtfReader assembly members.
    /// </summary>

    [Serializable]
    public class RtfReaderException : Exception
    {
        /// <summary>
        ///     Creates a new instance of RtfReaderException.
        /// </summary>
        public RtfReaderException() { }

        /// <summary>
        ///     Creates a new instance of RtfReaderException using the supplied message text.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        public RtfReaderException(string message) : base(message) { }

        /// <summary>
        ///     Creates a new instance of RtfReaderException using the supplied inner exception.
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException">
        ///     The exception that is the cause of the current exception, or a null reference
        ///     (Nothing in Visual Basic) if no inner exception is specified.
        /// </param>
        public RtfReaderException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        ///     Creates a new instance of RtfReaderException using the supplied serialized data.
        /// </summary>
        private RtfReaderException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}