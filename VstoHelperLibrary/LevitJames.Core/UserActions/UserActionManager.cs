// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Threading;
using JetBrains.Annotations;

namespace LevitJames.Core
{
    /// <summary>
    ///     A class for managing UserActions and optionally binding them to user interface elements.
    /// </summary>
    /// <remarks>
    ///     There are two types of user action a complex UserAction and a simple UserAction.
    ///     <para>
    ///         Complex user actions take the form of UserAction Types Complex user actions can be bound to user interface
    ///         elements when an appropriate UserActionAdapter is provided. Currently the following UserActionAdapters are
    ///         provided in the LJFramework.
    ///     </para>
    ///     <list>
    ///         <listheader>Provided UserActionAdapters</listheader>
    ///         <item>
    ///             RibbonUserActionAdapter (LevitJames.MSOffice Assembly) Allows binding to Microsoft Office Ribbon
    ///             elements. Note this class must bind to a RibbonUserAction Type rather than a regular UserAction Type.
    ///         </item>
    ///         <item>
    ///             CommandBarControlUserActionAdapter (LevitJames.MSOffice Assembly) Allows binding to Microsoft Office
    ///             CommandBarControl elements.
    ///         </item>
    ///         <item>
    ///             ButtonUserActionAdapter (LevitJames.WinForms Assembly) allows binding to controls which inherit from the
    ///             Microsoft WinForms ButtonBase Type (Buttons, Check boxes and Radio buttons).
    ///         </item>
    ///         <item>
    ///             ToolStripItemUserActionAdapter (LevitJames.WinForms Assembly). Allows binding to components which inherit
    ///             from the Microsoft WinForms ToolStripItem Type.
    ///         </item>
    ///     </list>
    ///     Adapters can be added and removed from the UserActionManager through the Adapters collection.
    ///     <para>
    ///         A simple user action takes the form of a simple string, and it's only purpose is to raise a
    ///         UserActionExecuted event using the id given. When a UserAction is executed a parameter can be supplied to
    ///         provide extra data for the UserAction.
    ///     </para>
    ///     <para>
    ///         When a UserAction is added to the UserActions collection it's Executed event is handled and forwarded through
    ///         to the UserActionManager.UserActionExecuted event. This provides a central location where UserActions can be
    ///         handled.
    ///     </para>
    /// </remarks>
    public static class UserActionManager
    {
        private static UserActionAdapterCollection _adapters; //Holds a collection of UserActionAdapter instances
        private static UserActionCollection _userActions; // Holds a collection of UserAction instances.

        private static List<UserActionTarget> _targets;
        private static UserActionTarget _target;
        private static UserActionTarget _globalTarget;
        private static Stack<string> _executionStack;


        /// <summary>
        ///     Returns/Set whether we raise events or not the UserActionManager can execute events
        /// </summary>
        
        public static bool Enabled { get; set; } = true;


        /// <summary>
        ///     The currently active target. This is where any executed UserActions are routed to for processing.
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public static IUserActionTarget ActiveTarget
        {
            set { _target = value != null ? GetTarget(value, throwIfnotFound: true) : null; }
            get { return _target?.Target; }
        }


        /// <summary>
        ///     Returns if there are any UserActions currently executing.
        /// </summary>
        public static string ExecutingUserAction
        {
            get
            {
                if (_executionStack == null || _executionStack.Count == 0)
                {
                    return string.Empty;
                }

                return _executionStack.Peek();
            }
        }


        private static Stack<string> ExecutionStack => _executionStack ?? (_executionStack = new Stack<string>());


        /// <summary>
        ///     Returns a collection of UserActionAdapter classes.
        /// </summary>
        
        /// <remarks>
        ///     UserActionAdapter classes are classes which convert the values in UserAction classes into there respective
        ///     User Interface equivalent.
        /// </remarks>
        public static UserActionAdapterCollection Adapters
            => _adapters ?? (_adapters = new UserActionAdapterCollection());


        /// <summary>
        ///     Returns a collection of UserAction Classes
        /// </summary>
        public static UserActionCollection UserActions => _userActions ?? (_userActions = new UserActionCollection());


        /// <summary>
        ///     A flag to signal if the Update method should be called to update the state of all the UserActions
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public static bool UpdateDirty { get; set; }

        /// <summary>
        ///     When set to true and a UserAction string property returns null, the UserAction will fire the UserActionGetString
        ///     event to retrieve the string from a resource.
        /// </summary>
        public static bool UseResourceStrings { get; set; }


        //Requirements

        //1. Get all the ui instances for a single UserAction.Id for the ActiveTarget
        //2. Get the UserAction for a single ui Instance
        //A UserAction can be bound to multiple UI instances
        //A single UserAction maps to multiple Instances
        //A single Instance maps to a single UserAction

        /// <summary>
        ///     Raised when a UserAction has been executed or when ExecuteUserAction was called with a string Id.
        /// </summary>
        public static event EventHandler<UserActionExecuteEventArgs> UserActionExecuting;

        /// <summary>
        ///     Raised when a UserAction has been executed or when ExecuteUserAction was called with a string Id.
        /// </summary>
        public static event EventHandler<UserActionExecuteEventArgs> UserActionExecuted;

        //// <summary>
        ////     Raised when a UserAction has been executed or when ExecuteUserAction was called with a string Id.
        //// </summary>
        //public static event EventHandler<UserActionLoadImageEventArgs> UserActionLoadImage;

        /// <summary>
        ///     Raised when a UserAction has been executed or when ExecuteUserAction was called with a string Id.
        /// </summary>
        public static event EventHandler<UserActionGetResourceEventArgs> UserActionGetResource;

        /// <summary>
        ///     Raised when a UserAction has been executed or when ExecuteUserAction was called with a string Id.
        /// </summary>
        public static event EventHandler<UserActionUpdateExceptionEventArgs> UserActionUpdateException;


        // ReSharper disable once ParameterOnlyUsedForPreconditionCheck.Local
        private static UserActionTarget GetTarget(IUserActionTarget target, bool throwIfnotFound = true)
        {
            if (_targets != null)
            {
                foreach (var t in _targets)
                    if (t.Target == target)
                        return t;
            }

            if (throwIfnotFound)
            {
                throw new LJException("Target not found");
            }

            return null;
        }


        /// <summary>
        ///     Adds a new target location where UserActions are routed too when executed.
        /// </summary>
        /// <param name="target"></param>
        public static void AddTarget(IUserActionTarget target)
        {
            if (GetTarget(target, throwIfnotFound: false) != null)
                throw new LJException("Target already exists");

            if (_targets == null)
                _targets = new List<UserActionTarget>();

            _targets.Add(new UserActionTarget(target));
        }


        /// <summary>
        ///     Removes a IUserActionTarget that was previously added via the AddTarget method
        /// </summary>
        /// <param name="target">The target to remove.</param>
        /// <returns>True if the target was successfully removed;false otherwise.</returns>
        public static bool RemoveTarget(IUserActionTarget target)
        {
            if (_targets == null)
                return false;

            var uaTarget = GetTarget(target, throwIfnotFound: false);

            if (_target != null && _target.Target == target)
                _target = null;

            if (uaTarget == null)
                return false;

            uaTarget.Clear();
            _targets.Remove(uaTarget);
            return true;
        }


        /// <summary>
        ///     Called from the UserActionCollection to provide central event handling of the UserAction.Executed Event.
        /// </summary>
        /// <param name="userAction">The UserAction Added.</param>
        internal static void OnUserActionAdded(UserAction userAction)
        {
            userAction.Executed -= OnUserActionExecute;
            userAction.Executed += OnUserActionExecute;
            userAction.PropertyChanged -= OnUserActionPropertyChanged;
            userAction.PropertyChanged += OnUserActionPropertyChanged;
        }


        /// <summary>
        ///     Called from the UserActionCollection to remove any UserAction Event handlers and any user interface elements bound
        ///     to the userAction to be removed.
        /// </summary>
        /// <param name="userAction">The UserAction removed.</param>
        internal static void OnUserActionRemoved(UserAction userAction)
        {
            userAction.Executed -= OnUserActionExecute;

            //Remove all the instances mapped to the user action

            _globalTarget?.UnbindFromUserAction(userAction);

            if (_targets == null)
                return;

            foreach (var target in _targets)
                target.UnbindFromUserAction(userAction);
        }


        private static bool TryUpdate(UserAction userAction, ref bool exception)
        {
            try
            {
                userAction.Update();
            }
            catch (Exception ex)
            {
                exception = true;
                if (UserActionUpdateException != null)
                {
                    if (OnUserActionExecute(new UserActionUpdateExceptionEventArgs(userAction, ex)) == false)
                        return false;
                }
                else
                {
                    throw;
                }
            }

            return true;
        }

        /// <summary>
        ///     Raises an UserActionExecuted Event using the id and parameter provided and returns True if the UserAction was
        ///     handled.
        /// </summary>
        /// <param name="id"></param>
        /// <param name="parameter">
        ///     Any value that may provide extra data for the executing UserAction. This value may be changed
        ///     by any event handlers of the the UserActionExecuted event. It's value is then returned from the call.
        /// </param>
        /// <returns>
        ///     The value of the parameter supplied or a different value if changed by any handlers of the UserActionExecuted
        ///     event.
        /// </returns>
        /// <remarks>
        ///     If the UserAction is not handled then an UserActionNotHandledException is thrown.
        ///     If UserActionManager.Enabled is False then no UserActions are executed, and no exception is raised.
        /// </remarks>
        public static bool TryExecuteUserAction([NotNull] string id, object parameter = null)
            => TryExecuteUserActionCore(id, parameter);


        internal static bool TryExecuteUserAction([NotNull] string id, bool autoCheck, object parameter = null)
            => TryExecuteUserActionCore(id, parameter, autoCheck);


        private static bool TryExecuteUserActionCore([NotNull] string id, object parameter, bool autoCheck = true)
        {
            if (Enabled == false)
                return false;

            ExecutionStack.Push(id);
            try
            {
                if (UserActions.TryGetValue(id, out UserAction userAction))
                {
                    if (!userAction.Enabled)
                        return true; // return true as the UserAction does exist, its just not enabled.

                    return userAction.ExecuteCore(autoCheck, parameter);
                }
                else
                {
                    var e = new UserActionExecuteEventArgs(id, parameter);
                    OnUserActionExecute(sender: null, e: e);

                    return e.Handled;
                }
            }
            finally
            {
                ExecutionStack.Pop();
            }
        }


        /// <summary>
        ///     Removes a UserAction binding from the supplied user interface element.
        /// </summary>
        /// <param name="uiElement"></param>
        public static bool UnbindFromUserAction([NotNull] object uiElement)
        {
            Check.NotNull(uiElement, nameof(uiElement));

            var adapter = GetAdapterFromUIElement(uiElement);
            if (adapter == null)
                return false;

            if (adapter.IsGlobal)
            {
                if (_globalTarget == null)
                    return false;

                _globalTarget.UnbindFromUserAction(adapter, uiElement);
            }
            else if (_target != null)
                return _target.UnbindFromUserAction(adapter, uiElement);

            return false;
        }


        /// <summary>
        ///     Occurs when a UserAction property has changed.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks>
        ///     This member forwards the property change to the corresponding adapter that will update the user interface
        ///     element(s) accordingly.
        /// </remarks>
        internal static void OnUserActionPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            var ua = (UserAction) sender;
            _globalTarget?.OnUserActionPropertyChanged(ua, e.PropertyName);
            _target?.OnUserActionPropertyChanged(ua, e.PropertyName);
        }


        /// <summary>
        ///     Occurs when a UserAction property has changed.
        /// </summary>
        /// <param name="e"></param>
        /// <remarks>
        ///     This member forwards the property change to the corresponding adapter that will update the user interface
        ///     element(s) accordingly.
        /// </remarks>
        [EditorBrowsable(EditorBrowsableState.Never)]
        public static void OnUserActionGetResource(UserActionGetResourceEventArgs e)
        {
            UserActionGetResource?.Invoke(sender: null, e: e);
        }


        ///// <summary>
        /////     Occurs when a UserAction property has changed.
        ///// </summary>
        ///// <param name="e"></param>
        ///// <remarks>
        /////     This member forwards the property change to the corresponding adapter that will update the user interface
        /////     element(s) accordingly.
        ///// </remarks>
        //[EditorBrowsable(EditorBrowsableState.Never)]
        //public static void OnUserActionLoadImage(UserActionLoadImageEventArgs e)
        //{
        //    UserActionLoadImage?.Invoke(sender: null, e: e);
        //}


        /// <summary>
        ///     Returns a UserActionAdapter which can bind to the user interface element supplied
        /// </summary>
        /// <param name="uiElement"></param>
        public static UserActionAdapter GetAdapterFromUIElement(object uiElement)
        {
            foreach (var adapter in Adapters)
            {
                if (adapter.CanBind(uiElement))
                    return adapter;
            }

            return null;
        }


        /// <summary>
        ///     Returns the UserAction that is bound to the passed uiElement
        /// </summary>
        /// <param name="uiElement">The UI Element such as a Ribbon string Windows Forms Button etc.</param>
        /// <returns>
        ///     A UserAction instance or null if either the ActiveTarget is null or the UI element is not bound to any
        ///     UserAction.
        /// </returns>
        public static UserAction UserActionFromUIElement(object uiElement)
        {
            if (_target == null)
                return null;

            var adapter = GetAdapterFromUIElement(uiElement);
            return adapter?.UserActionFromUIElement(uiElement);
        }


        /// <summary>
        ///     Returns an enumerator containing all the UI elements that are bound to the supplied UserAction
        /// </summary>
        /// <param name="userAction">
        ///     A UserAction instance to retrieve the bound UI elements for. The elements in the collection
        ///     are bound using the BindToUserAction methods.
        /// </param>
        public static IEnumerable UIElementsFromUserAction([NotNull] UserAction userAction)
        {
            if (_target == null)
                return null;

            var list = new ArrayList();
            list.AddRange(_target.GetUIElements(userAction));
            if (_globalTarget != null)
                list.AddRange(_globalTarget.GetUIElements(userAction));

            for (var i = list.Count - 1; i >= 0; i--)
            {
                var element = list[i];
                var adapter = GetAdapterFromUIElement(element);
                if (adapter == null)
                    continue;

                var unwrapped = adapter.GetUIElement(element, resolve: true);
                if (unwrapped != element)
                    list[i] = unwrapped;
            }

            return list;
        }


        /// <summary>
        ///     Calls Update only if UpdateDirty returns true.
        /// </summary>
        public static void UpdateIfDirty()
        {
            if (UpdateDirty)
                Update((string) null);
        }


        /// <summary>
        ///     Calls the Update method on all the UserAction classes contained in the UserActions Collection.
        /// </summary>
        public static void Update() => Update((string) null);

        /// <summary>
        ///     Calls the Update method on all the UserAction classes contained in the UserActions Collection using the filter
        ///     String provided.
        /// </summary>
        /// <param name="filter">
        ///     A String is used to match any UserAction.Id's that start with the same string.
        ///     <para>
        ///         For example passing a filter of "Application.Actions.Data" with find UserActions starting with an id of
        ///         "Application.Actions.Data" such as
        ///         "Application.Actions.Data.Copy" or "Application.Actions.Data.Cut"
        ///     </para>
        /// </param>
        /// <remarks>T</remarks>
        public static void Update(string filter)
        {
            if (_userActions == null)
                return;

            if (!Enabled)
            {
                UpdateDirty = true;
                return;
            }

            var anyExceptions = false;

            for (var index = 0; index < UserActions.Count; index++)
            {
                var userAction = UserActions[index];
                if (!string.IsNullOrEmpty(filter) &&
                    !userAction.Id.StartsWith(filter, StringComparison.OrdinalIgnoreCase))
                    continue;

                try
                {
                    userAction.Update();
                }
                catch (Exception ex)
                { 
                    if (UserActionUpdateException != null)
                    {
                        anyExceptions = true;
                        //We have event handler(s)
                        if (OnUserActionExecute(new UserActionUpdateExceptionEventArgs(userAction, ex)) == false)
                            break; //Don't continue looping'
                    }
                    else
                        throw;
                }
            }

            if (anyExceptions == false)
                //Only set to false if we had no exceptions
                UpdateDirty = false;
        }

        /// <summary>
        ///     Calls UserAction.Update for each UserAction.Id contained in userActions
        /// </summary>
        /// <param name="userActions">An IEnumerable containing the Id's of the UserActions to update.</param>
        /// <remarks>This method does not set UpdateDirty to False</remarks>
        public static void Update(IEnumerable<string> userActions)
        {
            if (_userActions == null)
                return;

            if (!Enabled)
            {
                UpdateDirty = true;
                return;
            }

            foreach (var ua in userActions)
            {
                if (string.IsNullOrEmpty(ua))
                    continue;

                var userAction = UserActions[ua];
                if (userAction == null)
                    continue;

                var exceptionThrown = false;
                if (TryUpdate(userAction, ref exceptionThrown) == false)
                    break;
            }
        }

        /// <summary>
        ///     Calls UserAction.Update for each UserAction contained in userActions
        /// </summary>
        /// <param name="userActions"></param>
        /// <remarks>This method does not set UpdateDirty to False</remarks>
        public static void Update(IEnumerable<UserAction> userActions)
        {
            if (_userActions == null || userActions == null)
                return;

            if (!Enabled)
            {
                UpdateDirty = true;
                return;
            }

            foreach (var userAction in userActions)
            {
                if (userAction == null)
                    continue;

                var exceptionThrown = false;
                if (TryUpdate(userAction, ref exceptionThrown) == false)
                    break;
            }
        }


        /// <summary>
        ///     Raises the UserActionUpdateException event if an exception is caught when calling UserActionManager.Update.
        /// </summary>
        private static bool OnUserActionExecute(UserActionUpdateExceptionEventArgs e)
        {
            UserActionUpdateException?.Invoke(sender: null, e: e);
            return e.Continue;
        }

        /// <summary>
        ///     Raises the UserActionExecuting event &amp; UserActionExecuted Events.
        /// </summary>
        /// <param name="sender">A UserAction instance.</param>
        /// <param name="e">The event arguments for the UserAction.Executed event.</param>
        /// <remarks>
        ///     The method catches all the UserAction.Executed events and re-raises them so they can be handled in a single
        ///     location.
        /// </remarks>
        private static void OnUserActionExecute(object sender, UserActionExecuteEventArgs e)
        {
            if (Enabled == false)
                return;

            UserActionExecuting?.Invoke(sender: null, e: e);
            if (!e.Handled)
                _target?.Target.UserActionExecuted(e);

            UserActionExecuted?.Invoke(sender: null, e: e);
        }


        /// <summary>
        ///     Raises an UserActionExecuted Event using the id and parameter provided and returns True if the UserAction was
        ///     handled..
        /// </summary>
        /// <param name="id"></param>
        /// <param name="parameter">
        ///     Any value that may provide extra data for the executing UserAction. This value may be changed
        ///     by any event handlers of the UserActionExecuted event. It's value is then returned from the call.
        /// </param>
        /// <returns>
        ///     The value of the parameter supplied or a different value if changed by any handlers of the UserActionExecuted
        ///     event.
        /// </returns>
        /// <remarks>
        ///     If the UserAction is not handled then an UserActionNotHandledException is thrown.
        ///     If UserActionManager.Enabled is False then no UserActions are executed, and no exception is raised.
        /// </remarks>
        public static object ExecuteUserAction([NotNull] string id, object parameter = null)
        {
            if (Enabled == false)
                return null;

            if (TryExecuteUserAction(id, parameter) == false)
                throw new UserActionNotHandledException(id);

            return parameter;
        }



        /// <summary>
        ///     Binds a UserAction to a user interface element.
        /// </summary>
        /// <param name="userAction">The UserAction which to associate with the user interface element</param>
        /// <param name="uiElement">The user interface element to associate with the UserAction</param>
        /// <remarks>
        ///     Binding a UserAction to a user interface element allows the user interface to automatically update when the
        ///     UserAction is up dated. Depending on the the UserInterfaceAdapter the UserAction's Execute method can be called
        ///     when the user interface element is acted upon. If there is no adapter available for the type of instance provided a
        ///     UserActionAdapterNotFoundException exception is thrown. An instance can only be added to a single UserAction.
        /// </remarks>
        [SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly",
            MessageId = "UserActionAdapter")]
        public static void BindToUserAction([NotNull] UserAction userAction, [NotNull] object uiElement)
        {
            Check.NotNull(userAction, nameof(userAction));
            Check.NotNull(uiElement, nameof(uiElement));

            var adapter = GetAdapterFromUIElement(uiElement);
            if (adapter == null)
                throw new UserActionAdapterNotFoundException("No UserActionAdapter found for the provided uiElement.");

            if (adapter.IsGlobal)
            {
                if (_globalTarget == null)
                    _globalTarget = new UserActionTarget();

                _globalTarget.BindToUserAction(adapter, userAction, uiElement);
            }
            else
            {
                if (_target == null && adapter.IsGlobal == false)
                    throw new InvalidOperationException("No Active Target");

                //Else try and Bind with a local Target Adapter
                _target?.BindToUserAction(adapter, userAction, uiElement);
            }
        }

        /// <summary>
        ///     Binds a UserAction to a user interface element.
        /// </summary>
        /// <param name="userAction">The UserAction which to associate with the user interface element</param>
        /// <param name="uiElements">An array of user interface element to associate with the UserAction</param>
        /// <remarks>
        ///     Binding a UserAction to a user interface element allows the user interface to automatically update when the
        ///     UserAction is up dated. Depending on the UserInterfaceAdapter the UserAction's Execute method can be called
        ///     when the user interface element is acted upon. If there is no adapter available for the type of instance provided a
        ///     UserActionAdapterNotFoundException exception is thrown. An instance can only be added to a single UserAction or a
        ///     UserActionInstanceAlreadyBound exception will be thrown.
        /// </remarks>
        public static void BindToUserAction([NotNull] UserAction userAction, [ItemNotNull] params object[] uiElements)
        {
            Check.NotNull(uiElements, "uiElements");
            foreach (var uiElement in uiElements)
                BindToUserAction(userAction, uiElement);
        }

        /// <summary>
        ///     Binds a UserAction to a user interface element.
        /// </summary>
        /// <param name="userActionId">
        ///     A string id representing the UserAction contained in the UserActions collection to associate
        ///     with the user interface element
        /// </param>
        /// <param name="uiElement">The user interface element to associate with the UserAction</param>
        /// <remarks>
        ///     Binding a UserAction to a user interface element allows the user interface to automatically update when the
        ///     UserAction is up dated. Depending on the UserInterfaceAdapter the UserAction's Execute method can be called
        ///     when the user interface element is acted upon. If there is no adapter available for the type of instance provided a
        ///     UserActionAdapterNotFoundException exception is thrown.
        /// </remarks>
        public static void BindToUserAction([NotNull] string userActionId, [NotNull] object uiElement)
            => BindToUserAction(UserActions[userActionId], uiElement);

        /// <summary>
        ///     Binds a UserAction to a user interface element.
        /// </summary>
        /// <param name="userActionId">
        ///     A string id representing the UserAction contained in the UserActions collection to associate
        ///     with the user interface element
        /// </param>
        /// <param name="uiElements">An array of user interface element to associate with the UserAction</param>
        /// <remarks>
        ///     Binding a UserAction to a user interface element allows the user interface to automatically update when the
        ///     UserAction is up dated. Depending on the UserInterfaceAdapter the UserAction's Execute method can be called
        ///     when the user interface element is acted upon. If there is no adapter available for the type of instance provided a
        ///     UserActionAdapterNotFoundException exception is thrown.
        /// </remarks>
        public static void BindToUserAction([NotNull] string userActionId, [NotNull] params object[] uiElements)
            => BindToUserAction(UserActions[userActionId], uiElements);

        /// <summary>
        ///     Binds a UserAction to a user interface element.
        /// </summary>
        /// <param name="userActionId">
        ///     A string id representing the UserAction contained in the UserActions collection to associate
        ///     with the user interface element
        /// </param>
        /// <param name="uiElements">A collection of user interface element to associate with the UserAction</param>
        /// <remarks>
        ///     Binding a UserAction to a user interface element allows the user interface to automatically update when the
        ///     UserAction is up dated. Depending on the UserInterfaceAdapter the UserAction's Execute method can be called
        ///     when the user interface element is acted upon. If there is no adapter available for the type of instance provided a
        ///     UserActionAdapterNotFoundException exception is thrown.
        /// </remarks>
        public static void BindToUserAction([NotNull] string userActionId, [NotNull] IEnumerable<object> uiElements)
        {
            Check.NotNull(uiElements, "uiElements");
            foreach (var uiElement in uiElements)
                BindToUserAction(UserActions[userActionId], uiElement);
        }

    }
}