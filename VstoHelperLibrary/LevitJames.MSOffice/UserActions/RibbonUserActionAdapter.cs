// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using LevitJames.MSOffice.MSWord;
using Microsoft.Office.Core;

namespace LevitJames.Core
{
    public class RibbonUserActionNotFoundEventArgs : EventArgs
    {
        public RibbonUserActionNotFoundEventArgs(string id, string tag, string propertyName)
        {
            Id = id;
            Tag = tag;
            PropertyName = propertyName;
        }

        public string Id { get; }
        public string Tag { get; }
        public string PropertyName { get; }
        public object Value { get; set; }
    }

    public class RibbonUserActionDropDownItem
    {

        // Properties
        public string Id { get; set; }

        public string Text { get; set; }

        public string ScreenTip { get; set; }

        public object Image { get; set; }

        public object Tag { get; set; }
    }

    /// <summary>
    ///     Provides support for binding UserActions to Microsoft Ribbon instances.
    /// </summary>
    /// <remarks>This class requires that the UserAction classes are of the derived type RibbonUserAction.</remarks>
    [Serializable]
    public sealed class RibbonUserActionAdapter : UserActionAdapter
    {
        public const string ContextMenuIdIdentifier = ".CM.";
        public const int ContextMenuIdIdentifierLength = 7; //".00.CM.";

        private bool _updateFromContextMenu;


        public override bool IsGlobal => true;
        public event EventHandler<RibbonUserActionNotFoundEventArgs> RibbonUserActionNotFound;


        protected override object Bind(UserAction userAction, object uiElement)
        {
            Check.NotNull(userAction, "userAction");

            if (!(userAction is RibbonUserAction))
            {
                throw new ArgumentException(@"Invalid UserAction, userAction must be of type RibbonUserAction.",
                                            nameof(userAction));
            }

            if (!(uiElement is string))
            {
                throw new ArgumentException(@"Invalid uiElement, uiElement must be of type String.", nameof(uiElement));
            }

            Check.NotEmpty((string)uiElement, "instance");

            return base.Bind(userAction, uiElement);
        }


        protected override void UpdateUIElement(UserAction userAction, object uiElement, bool binding)
        {
            Check.NotNull(userAction, nameof(userAction));
            Check.NotNull(uiElement, nameof(uiElement));
            if (!_updateFromContextMenu)
                WordExtensions.InvalidateRibbonControl(uiElement.ToString());
        }


        /// <summary>
        ///     Updates all  the ribbon controls to match the values in their respective UserActions.
        /// </summary>

        public static void UpdateAll()
        {
            if (UserActionManager.Enabled == false)
            {
                return;
            }

            WordExtensions.InvalidateRibbon();
        }


        private object OnUserActionNotFound(IRibbonControl control, string propertyName)
        {
            if (RibbonUserActionNotFound == null)
                return null;

            var e = new RibbonUserActionNotFoundEventArgs(control.Id, control.Tag, propertyName);
            RibbonUserActionNotFound.Invoke(this, e);
            return e.Value;
        }


        //Ribbon Control Callbacks


        /// <summary>
        ///     Called by Office and above, to notify when a ribbon button was clicked.
        /// </summary>
        /// <param name="control">The Ribbon control that executed the callback.</param>

        internal bool RibbonCallbackOnAction(IRibbonControl control) => RibbonCallbackOnAction(control.Id, control.Tag);


        internal bool RibbonCallbackOnAction(string id, string tag)
        {
            var rua = UserActionFromUIElement<RibbonUserAction>(id);
            if (rua == null)
                return false;

            object parameter = null;
            if (!string.IsNullOrEmpty(tag))
                parameter = tag;

            Execute(rua, parameter: parameter);
            return true;
        }

        /// <summary>
        ///     Called by Office and above, to notify when a ribbon button was clicked.
        /// </summary>
        /// <param name="control">The Ribbon control that executed the callback.</param>
        /// <param name="pressed"></param>

        internal bool RibbonCallbackToggleButtonOnAction(IRibbonControl control, bool pressed)
        {
            var rua = UserActionFromUIElement<RibbonUserAction>(control.Id);
            if (rua == null)
                return false;

            object parameter = null;
            if (!string.IsNullOrEmpty(control.Tag))
                parameter = control.Tag;

            rua.Checked = pressed;
            Execute(rua, parameter: parameter);
            return true;
        }


        /// <summary>
        ///     Called by Office and above, to notify when a ribbon button was clicked.
        /// </summary>
        /// <param name="control">The Ribbon control that executed the callback.</param>

        internal void RibbonCallbackOnChange(IRibbonControl control, string text)
        {
            var id = control.Id;
            var objA = UserActionFromUIElement<RibbonUserAction>(id);
            if (objA != null)
            {
                object parameter = null;
                if (!string.IsNullOrEmpty(text))
                {
                    parameter = text;
                }
                this.Execute(objA, parameter);
            }
        }



        /// <summary>
        ///     Called by Office 2007 and above, to request an image for a ribbon button.
        /// </summary>
        /// <param name="control">The Ribbon control that executed the callback.</param>
        /// <returns>An IDispPicture object containing the image for the button.</returns>

        [return: MarshalAs(UnmanagedType.IDispatch)]
        internal object RibbonCallbackGetButtonImage(IRibbonControl control)
        {
            var rua = UserActionFromUIElement<RibbonUserAction>(control.Id);

            if (rua != null)
            {
                //Debug.WriteLine("GetRibbonImage:" & control.Id)

                var size = control.Id.EndsWith(ContextMenuIdIdentifier, StringComparison.OrdinalIgnoreCase)
                               ? (int)UserActionImageSize.Regular
                               : GetImageSize(rua, control.Id, rua.DefaultImageSize);

                return rua.GetImage(size);

            }

            if (RibbonUserActionNotFound != null)
            {
                var ret = OnUserActionNotFound(control, nameof(UserAction.DefaultImageSize));
                var size = UserActionImageSize.Regular;
                if (ret != null)
                    size = (UserActionImageSize)ret;

                using (var img = (Image)OnUserActionNotFound(control, size == UserActionImageSize.Regular
                                                                           ? UserAction.RegularImage
                                                                           : UserAction.LargeImage))
                {
                    return img;
                }
            }

            return null;
        }


        internal bool RibbonCallbackGetEnabled(IRibbonControl control)
        {
            var rua = UserActionFromUIElement<RibbonUserAction>(control.Id);
            return rua?.EnabledResolved ??
                   Convert.ToBoolean(OnUserActionNotFound(control, nameof(UserAction.Enabled)), CultureInfo.InvariantCulture);

            //No UserAction found so True return enabled (ribbon may be checking a menu UserAction items that have not been created yet.
        }


        internal string RibbonCallbackGetLabel(IRibbonControl control)
        {
            var rua = UserActionFromUIElement<RibbonUserAction>(control.Id);
            return rua != null 
                ? rua.Text 
                : Convert.ToString(OnUserActionNotFound(control, nameof(UserAction.Text)), CultureInfo.InvariantCulture);
        }


        internal string RibbonCallbackGetEditValue(IRibbonControl control)
        {
            var rua = UserActionFromUIElement<RibbonUserAction>(control.Id);
            return rua != null 
                ? (string)rua.GetProperty("EditValue") 
                : Convert.ToString(OnUserActionNotFound(control, nameof(UserAction.Text)), CultureInfo.InvariantCulture);
        }


        internal int RibbonCallbackGetSize(IRibbonControl control)
        {
            var rua = UserActionFromUIElement<RibbonUserAction>(control.Id);
            if (rua != null)
            {
                if (control.Id.EndsWith(ContextMenuIdIdentifier, StringComparison.OrdinalIgnoreCase))
                    return (int)UserActionImageSize.Regular;

                return (int)GetImageSize(rua, control.Id, rua.DefaultImageSize);
            }

            return Convert.ToInt32(OnUserActionNotFound(control, nameof(UserAction.DefaultImageSize)), CultureInfo.InvariantCulture);
        }


        internal string RibbonCallbackGetScreenTip(IRibbonControl control)
        {
            var rua = UserActionFromUIElement<RibbonUserAction>(control.Id);
            return rua?.ToolTip;
        }


        internal string RibbonCallbackGetSuperTip(IRibbonControl control)
        {
            var rua = UserActionFromUIElement<RibbonUserAction>(control.Id);
            return rua?.SuperTip;
        }


        internal bool RibbonCallbackGetVisible(IRibbonControl control)
        {
            var id = control.Id; //For debugging
            var rua = UserActionFromUIElement<RibbonUserAction>(id);
            return rua?.VisibleResolved ??
                   Convert.ToBoolean(OnUserActionNotFound(control, nameof(UserAction.Visible)), CultureInfo.InvariantCulture);
            //No UserAction found so YES return True (ribbon may be checking a menu UserAction items that have not been created yet.
        }


        internal bool RibbonCallbackGetPressed(IRibbonControl control)
        {
            var rua = UserActionFromUIElement<RibbonUserAction>(control.Id);
            return rua != null && rua.Checked;
        }


        internal string RibbonCallbackGetKeyTip(IRibbonControl control)
        {
            var rua = UserActionFromUIElement<RibbonUserAction>(control.Id);
            return rua?.KeyTip;
        }


        internal string RibbonCallbackGetDescription(IRibbonControl control)
        {
            var rua = UserActionFromUIElement<RibbonUserAction>(control.Id);
            return rua?.Description;
        }
        internal int RibbonCallbackGetItemCount(IRibbonControl control)
        {

            var id = control.Id;
            var ua = UserActionFromUIElement<RibbonUserAction>(id);
            return ua?.DropDownItems?.Count ?? 0;

        }

        internal string RibbonCallbackGetComboText(IRibbonControl control)
        {
            var id = control.Id;
            var ua = UserActionFromUIElement<RibbonUserAction>(id);
            return ua?.SelectedItem?.Text;
        }

        internal string RibbonCallbackGetItemId(IRibbonControl control, int index)
        {
            var id = control.Id;
            var ua = UserActionFromUIElement<RibbonUserAction>(id);
            return ua?.DropDownItems?[index].Text;
        }

        [return: MarshalAs(UnmanagedType.IDispatch)]
        internal object RibbonCallbackGetItemImage(IRibbonControl control, int index)
        {
            var id = control.Id;
            var ua = UserActionFromUIElement<RibbonUserAction>(id);
            return ua?.DropDownItems?[index].Image;
        }

        internal string RibbonCallbackGetItemLabel(IRibbonControl control, int index)
        {
            var id = control.Id;
            var ua = UserActionFromUIElement<RibbonUserAction>(id);
            return ua?.DropDownItems?[index].Text;
        }
        internal string RibbonCallbackGetItemScreenTip(IRibbonControl control, int index)
        {
            var id = control.Id;
            var ua = UserActionFromUIElement<RibbonUserAction>(id);
            return ua?.DropDownItems?[index].ScreenTip;
        }
        internal int RibbonCallbackGetSelectedItemIndex(IRibbonControl control)
        {
            var id = control.Id;
            var ua = UserActionFromUIElement<RibbonUserAction>(id);
            return ua?.SelectedIndex ?? -1;
        }

        internal void RibbonCallbackDropDownAction(IRibbonControl control, string selectedItemId, int selectedIndex)
        {
            string id = control.Id;
            var ua = UserActionFromUIElement<RibbonUserAction>(id);
            if (ua == null)
                return;

            object parameter = null;
            if (!string.IsNullOrEmpty(selectedItemId))
                parameter = selectedItemId;

            ua.SelectedIndex = selectedIndex;
            Execute(ua, parameter);
        }

        ///// <summary>
        ///// Returns the key to use to identify the passed uiElement instance.
        ///// </summary>
        ///// <param name="uiElement">The instance to retrieve the key for.</param>
        ///// <returns>The default implementation uses the instance passed in as the key.</returns>
        //protected override object GetUIElementKey(object uiElement)
        //{
        //    var rc = (IRibbonControl)uiElement;
        //    return rc.Tag ?? rc.Id;
        //}

        /// <summary>
        ///     Returns the UserAction bound to the user interface element provided.
        /// </summary>
        /// <typeparam name="TUserAction">The Derived UserAction instance to return</typeparam>
        /// <param name="uiElement">The user interface element to return the UserAction for.</param>
        /// <returns>A UserAction class or null if no binding was found.</returns>
        public override TUserAction UserActionFromUIElement<TUserAction>(object uiElement)
        {

            var id = uiElement as string;
            if (string.IsNullOrEmpty(id))
                return null;

            var ua = base.UserActionFromUIElement<RibbonUserAction>(id);
            if (ua == null && id.EndsWith(ContextMenuIdIdentifier, StringComparison.OrdinalIgnoreCase))
            {
                _updateFromContextMenu = true;
                try
                {
                    var baseUserActionId = id.Substring(0, id.Length - ContextMenuIdIdentifierLength);
                    ua = base.UserActionFromUIElement<RibbonUserAction>(baseUserActionId);

                    if (ua != null)
                    {
                        Bind(ua, id); // Bind context menu Id to existing UserAction.
                    }
                }
                finally
                {
                    _updateFromContextMenu = false;
                }
            }

            return ua as TUserAction;
        }


        public override bool CanBind(object uiElement)
        {
            return uiElement is string;
        }

        public override bool CanBind(UserAction userAction, object uiElement)
        {
            return userAction is RibbonUserAction && uiElement is string;
        }


        protected override void OnUserActionChanged(UserAction userAction, object uiElement, string propertyName)
        {
            Check.NotNull(userAction, nameof(userAction));

            switch (propertyName)
            {
                //Standard user actions properties
                case nameof(UserAction.Text):
                case nameof(UserAction.DefaultImageSize):
                case nameof(UserAction.Enabled):
                case nameof(UserAction.CheckedState):
                case nameof(UserAction.RefreshImages):
                case nameof(UserAction.Visible):
                case UserAction.ParentEnabled:
                case UserAction.ParentVisible:

                case nameof(RibbonUserAction.Description):
                case nameof(RibbonUserAction.KeyTip):
                case nameof(RibbonUserAction.ScreenTip):
                case nameof(RibbonUserAction.ShowImage):
                case nameof(RibbonUserAction.ShowLabel):
                case nameof(RibbonUserAction.SuperTip):
                case nameof(RibbonUserAction.DropDownItems):
                    break;
                default:
                    var updateUi = userAction.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.CreateInstance | BindingFlags.FlattenHierarchy | BindingFlags.GetProperty)
                                             ?.GetCustomAttribute<UpdateUIAttribute>() != null;

                    if (updateUi)
                        break;
                    else
                        return;
            }

            UpdateUIElement(userAction, uiElement, binding: false);
        }
    }
}