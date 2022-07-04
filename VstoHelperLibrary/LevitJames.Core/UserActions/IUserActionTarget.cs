// © Copyright 2018 Levit & James, Inc.

using System.ComponentModel;

namespace LevitJames.Core
{
    /// <summary>
    ///     Defines the members for a class that is used as the Target for UserActions to be routed through too.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Advanced)]
    public interface IUserActionTarget
    {
        /// <summary>
        ///     Any object that provides context. The context is passed to the UserAction.Update delegate when the
        ///     UserAction.Update method is called.
        /// </summary>
        object Context { get; }

        /// <summary>
        ///     The method where an executed UserAction is routed through to when it's Execute method is called.
        /// </summary>
        /// <param name="e">
        ///     An instance of the UserActionExecuteEventArgs class that contains the details of the UserAction that
        ///     was executed.
        /// </param>
        void UserActionExecuted(UserActionExecuteEventArgs e);
    }
}