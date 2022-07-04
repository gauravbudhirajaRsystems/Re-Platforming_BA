// © Copyright 2018 Levit & James, Inc.

using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using JetBrains.Annotations;

namespace LevitJames.Core
{
    internal class UserActionTarget
    {
        private Dictionary<UserAction, ArrayList> _userActionUIElements;


        public UserActionTarget(IUserActionTarget target)
        {
            Target = target;
        }

        public UserActionTarget() { }

        public IUserActionTarget Target { get; }


        private Dictionary<UserAction, ArrayList> UserActionUIElements
            => _userActionUIElements ?? (_userActionUIElements = new Dictionary<UserAction, ArrayList>());

        public ICollection GetUIElements(UserAction userAction)
        {
            UserActionUIElements.TryGetValue(userAction, out var uiElements);
            return uiElements;
        }

        [SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId = "uiElement")]
        [SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly",
            MessageId = "UserActionAdapter")]
        public void BindToUserAction(UserActionAdapter adapter, UserAction userAction, object uiElement)
        {
            Check.NotNull(adapter, "adapter");

            // adapter.Bind may need to wrap the passed uiElement.
            // If it does then the wrapper class is returned and we need to use this class in our UIElements collection
            //If it is not wrapped the same uiElement is simply returned
            var boundUIElement = adapter.Bind(userAction, uiElement);
            if (boundUIElement == null)
                return;

            Debug.Assert(boundUIElement != null);

            var uiElements = (ArrayList) GetUIElements(userAction);
            if (uiElements == null)
            {
                uiElements = new ArrayList();
                UserActionUIElements.Add(userAction, uiElements);
            }

            uiElements.Add(boundUIElement);
        }


        public void Clear() => _userActionUIElements?.Clear();


        internal void OnUserActionPropertyChanged(UserAction userAction, string propertyName)
        {
            if (_userActionUIElements == null)
                return;

            UserActionAdapter adapter = null;
            var uiElements = GetUIElements(userAction);
            if (uiElements == null)
                return;

            foreach (var uiElement in uiElements)
            {
                if (adapter == null || adapter.CanBind(uiElement) == false)
                    adapter = UserActionManager.GetAdapterFromUIElement(uiElement);

                adapter.OnUserActionChanged(userAction, uiElement, propertyName);
            }
        }

        public bool UnbindFromUserAction([NotNull] UserAction userAction)
        {
            Check.NotNull(userAction, nameof(userAction));

            if (_userActionUIElements == null)
                return false;

            UserActionAdapter adapter = null;
            var uiElements = GetUIElements(userAction);
            if (uiElements == null)
                return false;

            foreach (var element in uiElements)
            {
                if (adapter == null || adapter.CanBind(element) == false)
                    adapter = UserActionManager.GetAdapterFromUIElement(element);

                adapter.Unbind(element);
            }

            return false;
        }

        public bool UnbindFromUserAction([NotNull] UserActionAdapter adapter, [NotNull] object uiElement)
        {
            Check.NotNull(adapter, nameof(adapter));
            Check.NotNull(uiElement, nameof(uiElement));

            if (_userActionUIElements == null)
                return false;

            var userAction = adapter.UserActionFromUIElement(uiElement);
            if (userAction == null)
                return false;

            var uiElements = (ArrayList) GetUIElements(userAction);
            if (uiElements == null)
                return false;

            for (var i = uiElements.Count - 1; i >= 0; i--)
            {
                var element = uiElements[i];
                if (!element.Equals(uiElement))
                    continue;

                adapter.Unbind(uiElement);
                uiElements.RemoveAt(i);
                return true;
            }

            return false;
        }
    }
}