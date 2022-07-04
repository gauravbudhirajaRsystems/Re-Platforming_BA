// © Copyright 2018 Levit & James, Inc.

using System;
using System.ComponentModel;
using System.Runtime.Serialization;

namespace LevitJames.Core
{
    /// <summary>
    ///     This exception is raised when a UserAction
    /// </summary>
    [Serializable]
    [EditorBrowsable(EditorBrowsableState.Advanced)]
    public sealed class UserActionAdapterNotFoundException : LJException
    {
        /// <summary>
        ///     Creates a new instance of UserActionAdapterNotFoundException
        /// </summary>
        public UserActionAdapterNotFoundException() { }

        /// <summary>
        ///     Creates a new instance of UserActionAdapterNotFoundException
        /// </summary>
        /// <param name="message">The message to show</param>
        public UserActionAdapterNotFoundException(string message) : base(message) { }

        /// <summary>
        ///     Creates a new instance of UserActionAdapterNotFoundException
        /// </summary>
        /// <param name="message">The message to show</param>
        /// <param name="innerException">The inner exception.</param>
        public UserActionAdapterNotFoundException(string message, Exception innerException)
            : base(message, innerException) { }

        /// <summary>
        ///     Creates a new instance of UserActionAdapterNotFoundException
        /// </summary>
        private UserActionAdapterNotFoundException(SerializationInfo info, StreamingContext context)
            : base(info, context) { }
    }
}