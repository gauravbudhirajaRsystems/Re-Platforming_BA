// © Copyright 2018 Levit & James, Inc.

using System;
using JetBrains.Annotations;

namespace LevitJames.Core
{
    /// <summary>
    ///     A class for storing the exception details of a userAction that failed when it was updating its state.
    /// </summary>
    public class UserActionUpdateExceptionEventArgs : EventArgs
    {
        /// <summary>
        ///     Creates a new instance of the UserActionUpdateExceptionEventArgs with the supplied arguments.
        /// </summary>
        /// <param name="userAction">The userAction that failed.</param>
        /// <param name="ex">The exception that was thrown.</param>
        public UserActionUpdateExceptionEventArgs([NotNull] UserAction userAction, [NotNull] Exception ex)
        {
            Check.NotNull(userAction, nameof(userAction));
            Check.NotNull(ex, nameof(ex));
            UserAction = userAction;
            Exception = ex;
        }

        /// <summary>
        ///     The Id of the UserAction that failed when updating its state
        /// </summary>
        public string Id => UserAction?.Id;

        /// <summary>
        ///     The UserAction that failed when updating its state.
        /// </summary>
        public UserAction UserAction { get; }

        /// <summary>
        ///     The exception that was thrown.
        /// </summary>
        public Exception Exception { get; }

        /// <summary>
        ///     return true to continue updating other UserActions/False to halt further processing.
        /// </summary>
        public bool Continue { get; set; } = true;
    }
}