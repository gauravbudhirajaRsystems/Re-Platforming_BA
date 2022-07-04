// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace LevitJames.Core
{
    /// <summary>
    ///     An exception to throw when an custom exception is thrown by the
    ///     LevitJames assemblies.
    /// </summary>
    [Serializable]
    public class LJException : Exception
    {
        /// <summary>
        ///     Creates a new instance of LJException
        /// </summary>
        public LJException() { }

        /// <summary>
        ///     Creates a new instance of LJException with the specified message
        /// </summary>
        /// <param name="message">The message to show</param>
        public LJException(string message) : base(message) { }

        /// <summary>
        ///     Creates a new instance of LJException with the specified message and inner exception
        /// </summary>
        /// <param name="message">The message to show</param>
        /// <param name="innerException">The inner exception.</param>
        public LJException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        ///     Creates a new instance of LJException from the supplied serialized data.
        /// </summary>
        /// <param name="info">
        ///     The <see cref="System.Runtime.Serialization.SerializationInfo" /> that holds the serialized object
        ///     data about the exception being thrown.
        /// </param>
        /// <param name="context">
        ///     The <see cref="System.Runtime.Serialization.StreamingContext" /> that contains contextual
        ///     information about the source or destination.
        /// </param>
        protected LJException(SerializationInfo info,
                              StreamingContext context) : base(info, context) { }

        /// <summary>
        ///     A list of associated files for diagnosing the cause of the exception.
        /// </summary>
        public List<string> DiagnosticFiles { get; set; }
    }
}