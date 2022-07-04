// © Copyright 2018 Levit & James, Inc.

using System;
using JetBrains.Annotations;

namespace LevitJames.Core
{
    /// <summary>
    ///     The event arguments used by the UserActionExecuted delegate
    /// </summary>
    public class UserActionEventArgs : EventArgs
    {
        /// <summary>
        ///     Creates a new instance of UserActionEventArgs
        /// </summary>
        protected UserActionEventArgs() { }

        /// <summary>
        ///     Creates a new UserActionExecuteEventArgs instance used to raise a simple user action.
        /// </summary>
        /// <param name="userAction">A complex UserAction instance.</param>
        public UserActionEventArgs([NotNull] UserAction userAction)
        {
            Check.NotNull(userAction, nameof(userAction));
            UserAction = userAction;
        }

        /// <summary>
        ///     A UserAction instance.
        /// </summary>
        
        /// <returns>If the UserAction raised is a complex type then the UserAction instance is returned, else a null.</returns>

        public UserAction UserAction { get; }
    }

    /// <summary>
    ///     The event arguments used by the UserActionExecuted delegate
    /// </summary>
    public class UserActionExecuteEventArgs : UserActionEventArgs
    {
        private readonly string _id;


        /// <summary>
        ///     Creates a new UserActionExecuteEventArgs instance used to raise a simple user action.
        /// </summary>
        /// <param name="id">A (preferably) unique string id to give the UserAction.</param>
        /// <param name="parameter">Any additional data to provide to the event. This value can be changed by the event handlers.</param>
        public UserActionExecuteEventArgs(string id, object parameter)
        {
            //If String.IsNullOrEmpty(id) Then
            //	Throw New ArgumentNullException("id")
            //End If
            _id = id;
            Parameter = parameter;
        }

        /// <summary>
        ///     Creates a new UserActionExecuteEventArgs instance used to raise a simple user action.
        /// </summary>
        /// <param name="userAction">A complex UserAction instance.</param>
        /// <param name="parameter">Any additional data to provide to the event. This value can be changed by the event handlers.</param>
        public UserActionExecuteEventArgs(UserAction userAction, object parameter) : base(userAction)
        {
            Parameter = parameter;
        }

        /// <summary>
        ///     Returns/sets if the user action has been handled.
        /// </summary>
        
        /// <returns>True if an event handler has handled the user action; false otherwise.</returns>

        public bool Handled { get; set; }

        /// <summary>
        ///     Returns\Sets and additional data useful for the user action.
        /// </summary>
        


        public object Parameter { get; set; }


        /// <summary>
        ///     The id of the user action executed. If the user action executed is a complex type then value is the UserAction.Id
        /// </summary>
        
        /// <returns>A string value representing the Id of the user action being executed.</returns>

        public string Id => UserAction != null ? UserAction.Id : _id;


        /// <summary>
        ///     Returns the short Name part of the UserAction.Id. So an Id of 'LevitJames.Editing.Bold' will return the name 'Bold'
        /// </summary>
        public string Name => UserAction != null ? UserAction.Name : UserAction.GetName(_id);
    }
}