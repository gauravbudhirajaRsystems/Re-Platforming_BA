// © Copyright 2018 Levit & James, Inc.

using System;
using System.Globalization;
using System.Runtime.Serialization;

namespace LevitJames.Core
{
    /// <summary>
    ///     Raised when a UserAction is executed but not handled by any targets.
    /// </summary>
    [Serializable]
    public sealed class UserActionNotHandledException : Exception
    {
        /// <summary>
        ///     Creates a new instance of UserActionNotHandledException
        /// </summary>
        /// <param name="userActionId">The string Id representing the UserAction.</param>
        public UserActionNotHandledException(string userActionId)
            : base(string.Format(CultureInfo.InvariantCulture, "UserAction {0} not handled", userActionId)) { }

        /// <summary>
        ///     Creates a new instance of UserActionNotHandledException
        /// </summary>
        public UserActionNotHandledException() { }

        /// <summary>
        ///     Creates a new instance of UserActionNotHandledException
        /// </summary>
        /// <param name="message">The error message that explains the reason for the exception.</param>
        /// <param name="innerException">The exception that is the cause of the current exception, or a null reference.</param>
        public UserActionNotHandledException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        ///     Initializes a new instance of the System.Exception class with serialized data.
        /// </summary>
        /// <param name="info">
        ///     The System.Runtime.Serialization.SerializationInfo that holds the serialized
        ///     object data about the exception being thrown.
        /// </param>
        /// <param name="context">
        ///     The System.Runtime.Serialization.StreamingContext that contains contextual information
        ///     about the source or destination.
        /// </param>
        private UserActionNotHandledException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}