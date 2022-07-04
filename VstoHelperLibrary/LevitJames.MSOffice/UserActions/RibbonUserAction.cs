// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace LevitJames.Core
{
    [AttributeUsage(AttributeTargets.Property)]
    public class UpdateUIAttribute : Attribute { }

    /// <summary>
    ///     Extends the UserAction class to provide support for Microsoft Office Ribbon controls.
    /// </summary>

    [Serializable]
    public class RibbonUserAction : UserAction
    {
        /// <summary>
        ///     Creates a UserAction instance using the provided string Id
        /// </summary>
        /// <param name="id">A unique string Id used to identify the UserAction.</param>
        /// <param name="text">The text to associate with the UserAction</param>
        /// <param name="image">the Image to associate with the UserAction.</param>
        /// <param name="toolTip">the ToolTip to associate with the UserAction.</param>
        /// <param name="enabled">sets if the UserAction is enabled or not.</param>
        /// <param name="visible">sets if the UserAction is visible or not.</param>
        /// <param name="keyTip">The string that represents the Key Combination to use to execute the UserAction.</param>
        /// <param name="largeImage">If the ribbon Size is Large then this is the parameter to supply a large Image.</param>
        /// <param name="superTip">A string value used to provide more detailed tooltip information.</param>
        /// <param name="description">
        ///     A string value to provide the descriptive text displayed under the text of a drop down menu
        ///     item.
        /// </param>
        /// <param name="updateDelegate"></param>
        /// <param name="defaultImageSize"></param>
        /// <remarks>The id cannot be a null or empty string.</remarks>
        [SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        public RibbonUserAction(string id, string text = null, object image = null, object largeImage = null,
                                string toolTip = null, string superTip = null, string keyTip = null, bool enabled = true,
                                bool visible = true, string description = null, UserActionUpdateDelegate updateDelegate = null,
                                UserActionImageSize defaultImageSize = UserActionImageSize.Large)
            : base(id, text, image, toolTip, enabled, visible, defaultImageSize, updateDelegate)
        {
            ShowImage = true;
            ShowLabel = true;
            if (!string.IsNullOrEmpty(keyTip))
            {
                SetPropertyValueCore(nameof(KeyTip), keyTip);
            }

            if (!string.IsNullOrEmpty(superTip))
            {
                SetPropertyValueCore(nameof(SuperTip), superTip);
            }

            if (!string.IsNullOrEmpty(description))
            {
                SetPropertyValueCore(nameof(Description), description);
            }

            if (largeImage != null)
            {
                SetPropertyValueCore(LargeImage, largeImage);
            }
        }


        public bool ShowImage
        {
            get => GetPropertyValue<bool>();
            set => SetPropertyValue(value);
        }


        public bool ShowLabel
        {
            get => GetPropertyValue<bool>();
            set => SetPropertyValue(value);
        }


        public string SuperTip
        {
            get => GetPropertyOrResourceString() ?? ToolTip;
            set => SetPropertyValue(value);
        }


        public string ScreenTip
        {
            get => GetPropertyOrResourceString() ?? ToolTip;
            set => SetPropertyValue(value);
        }


        public string KeyTip
        {
            get => GetPropertyOrResourceString(); // ?? ToolTip;
            set => SetPropertyValue(value);
        }


        public string Description
        {
            get => GetPropertyOrResourceString() ?? ToolTip;
            set => SetPropertyValue(value);
        }

        /// <summary>
        /// Creates a new DropDownItems collection containing the items supplied.
        /// </summary>
        /// <param name="items">A collection of EnumItem values to add. The existing collection is cleared when calling this method.</param>
        public void AddDropDownItems(IEnumerable<TextValuePair<Enum>> items)
        {
            Check.NotNull(items, nameof(items));

            var remove = items == null || !items.Any();
            if (!GetDropDownItems(remove, out var dropDownItems))
                return;

            if (remove)
                return;

            foreach (var item in items)
            {

                var dropDownItem = new RibbonUserActionDropDownItem();
                new RibbonUserActionDropDownItem().Id = Id + "_" + item.Value;

                dropDownItem.Text = item.Text;
                dropDownItem.Tag = item.Value;
                dropDownItems.Add(dropDownItem);
            }

            OnPropertyChanged(nameof(DropDownItems));
        }

        /// <summary>
        /// Creates a new DropDownItems collection containing the items supplied.
        /// </summary>
        /// <param name="items">A collection of RibbonUserActionDropDownItem values to add. The existing collection is cleared when calling this method.</param>
        public void AddDropDownItems(IEnumerable<RibbonUserActionDropDownItem> items)
        {
            Check.NotNull(items, nameof(items));

            var remove = items == null || !items.Any();
            if (!GetDropDownItems(remove, out var dropDownItems))
                return;

            if (remove)
                return;

            foreach (var item in items)
            {
                if (string.IsNullOrEmpty(item.Id))
                    item.Id += "_" + dropDownItems.Count;

                dropDownItems.Add(item);
            }

            OnPropertyChanged(nameof(DropDownItems));

        }

        public IReadOnlyList<RibbonUserActionDropDownItem> DropDownItems => GetPropertyValue<IReadOnlyList<RibbonUserActionDropDownItem>>();

        private bool GetDropDownItems(bool remove, out IList<RibbonUserActionDropDownItem> items)
        {

            items = (IList<RibbonUserActionDropDownItem>)DropDownItems;
            if (items != null)
            {
                if (remove)
                {
                    SetPropertyValue<object>(null, raisePropertyChange: true);
                    return false;
                }
            }
            else if (remove)
                return true;

            items = items ?? new List<RibbonUserActionDropDownItem>();
            items.Clear();

            SetPropertyValue(items, nameof(DropDownItems), false); // Will raise event after items have been added.

            return true;
        }

        public int SelectedIndex
        {
            get =>
                GetPropertyValue<int>();
            set
            {
                if (SetPropertyValue(value, nameof(SelectedIndex), true))
                    OnPropertyChanged(nameof(SelectedIndex));
            }
        }

        public RibbonUserActionDropDownItem SelectedItem
        {
            get
            {
                RibbonUserActionDropDownItem item;
                int selectedIndex = this.SelectedIndex;
                if (selectedIndex == -1)
                    return null;

                var dropDownItems = (IList<RibbonUserActionDropDownItem>)this.DropDownItems;
                return ((dropDownItems != null) && (selectedIndex < dropDownItems.Count)) ? dropDownItems[selectedIndex] : null;

            }
            set
            {
                var dropDownItems = (IList<RibbonUserActionDropDownItem>)this.DropDownItems;
                SelectedIndex = !((dropDownItems == null) || ReferenceEquals(value, null)) ? dropDownItems.IndexOf(value) : -1;
            }
        }

        public object ImageResolved
        {
            get
            {
                if (ShowImage == false)
                    return null;

                if (DefaultImageSize == UserActionImageSize.Large)
                    return GetImage(UserActionImageSize.Large);

                return GetImage(UserActionImageSize.Regular);
            }
        }


        protected override bool GetBuiltInProperty(string name, out object value)
        {
            switch (name)
            {
                case nameof(ShowImage):
                    value = ShowImage;
                    break;
                case nameof(ShowLabel):
                    value = ShowLabel;
                    break;
                case nameof(SuperTip):
                    value = ShowImage;
                    break;
                case nameof(KeyTip):
                    value = GetPropertyValue<string>(nameof(KeyTip));
                    break;
                case nameof(Description):
                    value = GetPropertyValue<string>(nameof(Description));
                    break;
                case nameof(ScreenTip):
                    value = GetPropertyValue<string>(nameof(ScreenTip));
                    break;
                case nameof(DropDownItems):
                    value = GetPropertyValue<IReadOnlyList<RibbonUserActionDropDownItem>>(nameof(DropDownItems));
                    break;
                case nameof(SelectedIndex):
                    value = GetPropertyValue<int>(nameof(SelectedIndex));
                    break;
                default:
                    return base.GetBuiltInProperty(name, out value);
            }

            return false;
        }


        protected override bool SetBuiltInProperty(string name, object value)
        {
            switch (name)
            {
                case nameof(ShowImage):
                    ShowImage = Convert.ToBoolean(value);
                    break;
                case nameof(ShowLabel):
                    ShowLabel = Convert.ToBoolean(value);
                    break;
                case nameof(SuperTip):
                    ShowImage = Convert.ToBoolean(value);
                    break;
                case nameof(KeyTip):
                    KeyTip = Convert.ToString(value);
                    break;
                case nameof(Description):
                    Description = Convert.ToString(value);
                    break;
                case nameof(ScreenTip):
                    ScreenTip = Convert.ToString(value);
                    //Case LargeImageProperty : Me.LargeImage = DirectCast(value, System.Drawing.Image)
                    break;
                case nameof(SelectedIndex):
                    SelectedIndex = Convert.ToInt32(value);
                    break;
                default:
                    return base.SetBuiltInProperty(name, value);
            }

            return false;
        }



    }
}