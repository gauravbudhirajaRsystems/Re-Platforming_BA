// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using JetBrains.Annotations;

namespace LevitJames.Core
{
    /// <summary>
    ///     A base class the provides the base functionality for converting UserAction properties into there respective user
    ///     interface types.
    /// </summary>
    /// <remarks>
    ///     <list>
    ///         <listheader>Default derived classes are:</listheader>
    ///         <item>LevitJames.UserActions.UserActionGroupAdapter</item>
    ///         <item>LevitJames.UserActions.ButtonUserAction (LevitJames.WinForms assembly)</item>
    ///         <item>LevitJames.UserActions.ToolStripItem (LevitJames.WinForms assembly)</item>
    ///         <item>LevitJames.UserActions.CommandBarControlUserActionAdapter (LevitJames.MSOffice assembly)</item>
    ///         <item>LevitJames.UserActions.RibbonUserActionAdapter  (LevitJames.MSOffice assembly)</item>
    ///     </list>
    ///     <para></para>
    ///     This class inherits from MarshalByRefObject so that it can receive calls across AppDomains.
    /// </remarks>
    [Serializable]
    [EditorBrowsable(EditorBrowsableState.Advanced)]
    public abstract class UserActionAdapter
    {
        private readonly object _syncLock = new object();

        private Dictionary<object, UserAction> _uiElements;
        private bool _useUiElementImage;


        /// <summary>
        ///     Returns a Dictionary collection of UserActions, There key is the user interface element.
        /// </summary>
        


        private Dictionary<object, UserAction> UIElementsInternal => _uiElements ?? (_uiElements = new Dictionary<object, UserAction>());


        /// <summary>
        ///     Returns if the UserActionAdapter is shared between all UserActionTargets or if each UserActionTarget contains its
        ///     own UserActionAdapter instance.
        /// </summary>
        public virtual bool IsGlobal => false;


        /// <summary>
        ///     Returns Sets if the Image is Retrieved from the UserAction.Image property or the UI Element when binding
        /// </summary>
        /// <value>
        ///     True to use the UI Element's Image when UserActionManager.BindToUserAction is called; false to use the image
        ///     supplied by the UserAction.Image property.
        /// </value>
        /// <remarks>This property is only used when UserActionManager.BindToUserAction is called.</remarks>
        public bool UseUIElementImageWhenBinding
        {
            get { return _useUiElementImage; }
            set { _useUiElementImage = value; }
        }


        /// <summary>
        ///     Called by the UserInterfaceManager when a UserAction property has changed.
        /// </summary>
        /// <param name="userAction"></param>
        /// <param name="uiElement"></param>
        /// <param name="propertyName"></param>
        /// <remarks>It is in this member that user interface elements update to reflect any changes in the UserAction.</remarks>
        protected internal abstract void OnUserActionChanged(UserAction userAction, object uiElement,
                                                             string propertyName);

        /// <summary>
        ///     Returns whether this adapter can bind the supplied userAction to the given uiElement
        /// </summary>
        /// <param name="userAction">The UserAction to bind to. Cannot be null</param>
        /// <param name="uiElement">A user interface element to bind to. Cannot be null.</param>
        /// <returns>True if this adapter supports binding the UserAction to the supplied uiElement; false otherwise.</returns>
        public abstract bool CanBind(UserAction userAction, object uiElement);

        /// <summary>
        ///     Returns whether this adapter can bind the supplied userAction to the given uiElement
        /// </summary>
        /// <param name="uiElement">A user interface element to bind to. Cannot be null.</param>
        /// <returns>True if this adapter supports binding the UserAction to the supplied uiElement; false otherwise.</returns>
        public abstract bool CanBind(object uiElement);


        /// <summary>
        ///     Returns the key to use to identify the passed uiElement instance.
        /// </summary>
        /// <param name="uiElement">The instance to retrieve the key for.</param>
        /// <returns>The default implementation uses the instance passed in as the key.</returns>
        protected virtual object GetUIElementKey([NotNull] object uiElement) => uiElement;


        /// <summary>
        ///     Called by the UserActionManager when a new instance is to be added to the class.
        /// </summary>
        /// <param name="userAction">The UserAction to bind to</param>
        /// <param name="uiElement">
        ///     The user interface element to bind to. A user interface element can only bind to a single
        ///     UserAction.
        /// </param>
        protected internal virtual object Bind([NotNull] UserAction userAction, [NotNull] object uiElement)
        {
            Check.NotNull(userAction, nameof(userAction));
            Check.NotNull(uiElement, nameof(uiElement));

            if (UIElementsInternal.ContainsKey(GetUIElementKey(uiElement)))
            {
                return uiElement;
                //Throw New UserActionInstanceAlreadyBoundException("The instance supplied is already bound to a UserAction.")
            }

            var uiElementCore = AddCore(userAction, uiElement);

            UpdateUIElement(userAction, uiElement, binding: true);

            return uiElementCore;
        }


        /// <summary>
        ///     Adds the passed uiElement to this adapters collection of UIElements.
        /// </summary>
        /// <param name="userAction">The UserAction instance associated with the uiElement.</param>
        /// <param name="uiElement">The user interface element to associate with the UserAction</param>
        protected virtual object AddCore([NotNull] UserAction userAction, [NotNull] object uiElement)
        {
            Check.NotNull(userAction, nameof(userAction));
            Check.NotNull(uiElement, nameof(uiElement));

            lock (_syncLock)
            {
                UIElementsInternal.Add(GetUIElementKey(uiElement), userAction);
            }

            return uiElement;
        }


        /// <summary>
        ///     Called by the UserActionManager when a user interface element is to be removed from the class.
        /// </summary>
        /// <param name="uiElement">The user interface element to unbind (remove) from this adapter.</param>
        /// TThe user interface element to un bind.
        protected internal virtual void Unbind([NotNull] object uiElement)
        {
            //If This method is changed the Clear method may also need altering.
            Check.NotNull(uiElement, nameof(uiElement));

            if (_uiElements == null)
                return;

            lock (_syncLock)
                UIElementsInternal.Remove(GetUIElementKey(uiElement));
        }


        /// <summary>
        ///     Called by the UserActionManager when all the user interface elements must be cleared, such as when the adapter is
        ///     removed from the UserActionManager
        /// </summary>
        protected internal virtual void Clear()
        {
            if (_uiElements == null)
                return;

            lock (_syncLock)
                UIElementsInternal.Clear();
        }


        /// <summary>
        ///     Returns the size of image
        /// </summary>
        /// <param name="userAction"></param>
        /// <param name="uiElement"></param>
        /// <param name="defaultSize"></param>
        protected virtual UserActionImageSize GetImageSize([NotNull] UserAction userAction, [NotNull] object uiElement,
                                                           UserActionImageSize defaultSize)
        {
            Check.NotNull(userAction, nameof(userAction));

            var propName = GetType().Name + "." + nameof(UserAction.DefaultImageSize);
            if (!userAction.HasProperty(propName))
                return defaultSize;

            return (UserActionImageSize) userAction.GetProperty(propName);
        }


        /// <summary>
        ///     Returns if the UserAction supports setting the Image property with the size of image supplied.
        /// </summary>
        /// <param name="imageSize">The size of the image to be set.</param>
        /// <param name="propertyName"></param>
        protected static bool IsImagePropertySettable(UserActionImageSize imageSize, string propertyName)
        {
            if (string.IsNullOrEmpty(propertyName) == false)
            {
                if (propertyName == UserAction.RegularImage)
                {
                    if (imageSize != UserActionImageSize.Regular)
                        return false;
                }
                else if (propertyName == UserAction.LargeImage)
                {
                    if (imageSize != UserActionImageSize.Large)
                        return false;
                }
            }

            return true;
        }


        /// <summary>
        ///     Override to update the passed uiElement with the values from the UserAction.
        /// </summary>
        /// <remarks>Default implementation does nothing</remarks>
        protected virtual void UpdateUIElement(UserAction userAction, object uiElement, bool binding) { }


        /// <summary>
        ///     The handler used when a UI element contains a Disposed event. When hooked up in the Bind call if the UIElement is
        ///     still bound to a UserAction.
        ///     This handler will Automatically unbind the UserAction when the control is disposed.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected virtual void UIElementDisposedHandler(object sender, EventArgs e)
        {
            UserActionManager.UnbindFromUserAction(sender);
        }


        /// <summary>
        ///     Called when the UI element performs some action. For example a button click or an item is selected from a combo box
        ///     etc.
        /// </summary>
        /// <param name="userAction">The UserAction associated with the UI element that will be executed.</param>
        /// <param name="parameter">A parameter value that can be passed to and from the UserAction.</param>
        [SuppressMessage("Microsoft.Design", "CA1007:UseGenericsWhereAppropriate")]
        [SuppressMessage("Microsoft.Design", "CA1045:DoNotPassTypesByReference", MessageId = "1#")]
        protected virtual bool Execute([NotNull] UserAction userAction, object parameter = null)
        {
            Check.NotNull(userAction, nameof(userAction));
            return UserActionManager.TryExecuteUserAction(userAction.Id, userAction.AutoCheck, parameter);
        }


        /// <summary>
        ///     Returns the Bound UIElement that is conically the same as the uiElement passed in.
        ///     This is required for Microsoft Office CommandBarControl support as, command bar elements can use different
        ///     instances to represent the same control over time.
        /// </summary>
        /// <param name="uiElement">The internally bound UIElement to return</param>
        protected internal virtual object GetUIElement(object uiElement) => GetUIElement(uiElement, resolve: false);

        /// <summary>
        ///     Returns the Bound UIElement that is conically the same as the uiElement passed in.
        ///     This is required for Microsoft Office CommandBarControl support as, command bar elements can use different
        ///     instances to represent the same control over time.
        /// </summary>
        /// <param name="uiElement">The internally bound UIElement to return</param>
        /// <param name="resolve"></param>
        protected internal virtual object GetUIElement(object uiElement, bool resolve) => uiElement;


        /// <summary>
        ///     Returns an IEnumerable containing a KeyValuePair of Object (UIElement) and it's bound UserAction
        /// </summary>
        [SuppressMessage("Microsoft.Design", "CA1006:DoNotNestGenericTypesInMemberSignatures")]
        protected IEnumerable<KeyValuePair<object, UserAction>> GetItems() => _uiElements;


        /// <summary>
        ///     Returns the UserAction bound to the user interface element provided.
        /// </summary>
        /// <param name="uiElement">The user interface element to return the UserAction for.</param>
        /// <returns>A UserAction class or null if no binding was found.</returns>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public UserAction UserActionFromUIElement(object uiElement) => UserActionFromUIElement<UserAction>(uiElement);

        /// <summary>
        ///     Returns the UserAction bound to the user interface element provided.
        /// </summary>
        /// <typeparam name="TUserAction">The Derived UserAction instance to return</typeparam>
        /// <param name="uiElement">The user interface element to return the UserAction for.</param>
        /// <returns>A UserAction class or null if no binding was found.</returns>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public virtual TUserAction UserActionFromUIElement<TUserAction>(object uiElement) where TUserAction : UserAction
        {
            UserAction userAction;
            UIElementsInternal.TryGetValue(GetUIElementKey(uiElement), out userAction);
            return userAction as TUserAction;
        }
    }
}