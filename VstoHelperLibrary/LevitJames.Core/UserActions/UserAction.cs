// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

namespace LevitJames.Core
{
    /// <summary>
    ///     A class that represents the basic properties for a user performed action, including information to display to the
    ///     user.
    /// </summary>
    [Serializable]
    public class UserAction : INotifyPropertyChanged
    {
        ///<Summary>The string constant used for storing the State Property value</Summary>
        private const string StateProperty = "_$state$_";

        /// <summary>
        ///     The property name used for storing the regular image value.
        /// </summary>
        public const string RegularImage = "Image.Regular";

        /// <summary>
        ///     The property name used for storing the large image value.
        /// </summary>
        public const string LargeImage = "Image.Large";

        /// <summary>
        ///     The property name used for storing if a parent UserAction is visible
        /// </summary>
        public const string ParentVisible = "Parent.Visible";

        /// <summary>
        ///     The property name used for storing if a parent UserAction is enabled
        /// </summary>
        public const string ParentEnabled = "Parent.Enabled";

        [NonSerialized] private readonly UserActionUpdateDelegate _updateDelegate;


        // Don't Serialize events or delegates as the serializer will the serialize the sink instance as well.
        [NonSerialized] private EventHandler<UserActionExecuteEventArgs> _executed;

        private ListDictionary _properties; // Dictionary<string, object> _properties;

        [NonSerialized] private PropertyChangedEventHandler _propertyChanged;


        /// <summary>
        ///     Creates a UserAction instance using the provided string Id
        /// </summary>
        /// <param name="id">A unique string Id used to identify the UserAction.</param>
        /// <param name="text">The text to associate with the UserAction</param>
        /// <param name="image">The Image to associate with the UserAction .</param>
        /// <param name="toolTip">The ToolTip text to associate with the UserAction . </param>
        /// <param name="enabled">The enabled state of the UserAction</param>
        /// <param name="visible">The visibility of the UserAction</param>
        /// <param name="defaultImageSize"></param>
        /// <param name="updateDelegate">A callback used when the Update member of the UserAction is called.</param>
        /// <remarks>The id cannot be a null or empty string.</remarks>
        public UserAction(string id, string text = null, object image = null, string toolTip = null, bool enabled = true,
                          bool visible = true, UserActionImageSize defaultImageSize = UserActionImageSize.Regular,
                          UserActionUpdateDelegate updateDelegate = null)
        {
            if (string.IsNullOrEmpty(id))
                throw new ArgumentNullException(nameof(id));

            //_properties = new Dictionary<string, object>();
            _properties = new ListDictionary();

            Id = string.Intern(id);

            if (!string.IsNullOrEmpty(text))
                SetPropertyValueCore(nameof(Text), text);

            if (!string.IsNullOrEmpty(toolTip))
                SetPropertyValueCore(nameof(ToolTip), toolTip);

            if (image != null)
                SetPropertyValueCore(RegularImage, image);

            UserActionStates state = 0;
            if (visible)
                state |= UserActionStates.Visible | UserActionStates.ParentVisible;
            else
                state |= UserActionStates.ParentVisible;

            if (enabled)
                state |= UserActionStates.Enabled | UserActionStates.ParentEnabled;
            else
                state |= UserActionStates.ParentEnabled;

            if (defaultImageSize == UserActionImageSize.Large)
                state |= UserActionStates.ImageSizeLarge;

            SetPropertyValueCore(StateProperty, state);
            _updateDelegate = updateDelegate;
        }


        /// <summary>
        ///     The Unique Id to give the UserAction.
        /// </summary>



        public string Id { get; }


        /// <summary>
        ///     Returns the last part of a UserAction's Id
        /// </summary>

        /// <remarks>
        ///     For a UserAction with an Id of "SomeCompany.SomeApplication.Close" this member will return the last string
        ///     element. In this example this is would be "Close"
        /// </remarks>
        public string Name => GetName(Id);


        /// <summary>
        ///     Returns/sets the text that is displayed to the user.
        /// </summary>

        /// <remarks>Raises the PropertyChanged event when the value changes.</remarks>
        public virtual string Text
        {
            get => GetPropertyOrResourceString();
            set => SetPropertyValue(value);
        }


        /// <summary>
        ///     Returns/sets if the UserAction is enabled or not.
        /// </summary>

        /// <remarks>
        ///     Raises the PropertyChanged event when the value changes. The UserAction.Execute methods cannot be called when
        ///     Enabled = False, or an UserActionNotAvailable will be thrown.
        /// </remarks>
        public virtual bool Enabled
        {
            get => GetState(UserActionStates.Enabled);
            set => SetState(UserActionStates.Enabled, value);
        }


        /// <summary>
        ///     Returns if this UserAction is Enabled. If the UserAction is owned by a Parent UserAction then the Enabled state of
        ///     the Parent UserAction is resolved along with the Enabled state of this instance.
        ///     Typically this is set automatically when using the UserActionGroupAdapter, So that you can set the enabled state of
        ///     the group and all the child UserActions automatically enable/disable.
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public bool EnabledResolved => GetState(UserActionStates.Enabled | UserActionStates.ParentEnabled);


        /// <summary>
        ///     Returns/sets if the UserAction is enabled or not.
        /// </summary>

        /// <remarks>Raises the PropertyChanged event when the value changes.</remarks>
        public virtual bool Visible
        {
            get => GetState(UserActionStates.Visible);
            set => SetState(UserActionStates.Visible, value);
        }


        /// <summary>
        ///     Returns if this UserAction is Visible. If the UserAction is owned by a Parent UserAction then the Visible state of
        ///     the Parent UserAction is resolved along with the Visible state of this instance.
        ///     Typically this is set automatically when using the UserActionGroupAdapter, So that you can set the visibility of
        ///     the group and all the child UserActions automatically show/hide.
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public bool VisibleResolved => GetState(UserActionStates.Visible | UserActionStates.ParentVisible);


        /// <summary>
        ///     Returns/sets if the UserAction is checked or not.
        /// </summary>

        /// <remarks>Raises the PropertyChanged event when the value changes.</remarks>
        public bool Checked
        {
            get => CheckedState == UserActionCheckState.Checked;
            set => CheckedState = value ? UserActionCheckState.Checked : UserActionCheckState.Unchecked;
        }


        /// <summary>
        ///     Returns/sets the checked state of the UserAction.
        /// </summary>

        /// <remarks>Raises the PropertyChanged event when the value changes.</remarks>
        public virtual UserActionCheckState CheckedState
        {
            get
            {
                if (GetState(UserActionStates.Checked))
                    return UserActionCheckState.Checked;

                return GetState(UserActionStates.Indeterminate)
                           ? UserActionCheckState.Indeterminate
                           : UserActionCheckState.Unchecked;
            }
            set
            {
                var checkedState = CheckedState;
                if (checkedState == value)
                    return;
                var state = (UserActionStates)GetProperty(StateProperty);

                switch (value)
                {
                    case UserActionCheckState.Checked:
                        state |= UserActionStates.Checked;
                        state &= ~UserActionStates.Indeterminate;

                        break;
                    case UserActionCheckState.Indeterminate:
                        state &= ~UserActionStates.Checked;
                        state |= UserActionStates.Indeterminate;

                        break;
                    case UserActionCheckState.Unchecked:
                        state &= ~(UserActionStates.Checked | UserActionStates.Indeterminate);
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(value), value, message: null);
                }

                // ReSharper disable once ExplicitCallerInfoArgument
                if (SetPropertyValue(state, StateProperty, raisePropertyChange: false))
                    OnPropertyChanged(nameof(CheckedState));
            }
        }


        /// <summary>
        ///     Returns if the UI element for this UserAction should automatically check its state when clicked.
        /// </summary>
        public bool AutoCheck
        {
            get => GetState(UserActionStates.AutoCheck);
            set => SetState(UserActionStates.AutoCheck, value);
        }


        /// <summary>
        ///     Returns if the UserAction is Executing. This can help detect recursive calls.
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public bool Executing => GetState(UserActionStates.Executing);


        /// <summary>
        ///     The default image size to use for the UIElement.
        /// </summary>
        public UserActionImageSize DefaultImageSize
        {
            get => GetState(UserActionStates.ImageSizeLarge)
                       ? UserActionImageSize.Large
                       : UserActionImageSize.Regular;
            set => SetState(UserActionStates.ImageSizeLarge, value == UserActionImageSize.Large);
        }


        /// <summary>
        ///     A delegate used to update the State of UserAction when the Update methods are called.
        /// </summary>
        protected UserActionUpdateDelegate UpdateDelegate => _updateDelegate;


        /// <summary>
        ///     Returns/sets the ToolTip text for the UserAction .
        /// </summary>

        /// <remarks>Raises the PropertyChanged event when the value changes.</remarks>
        public string ToolTip
        {
            get
            {
                var value = GetPropertyOrResourceString();
                return string.IsNullOrEmpty(value) ? Text : value;
            }
            set => SetPropertyValue(value);
        }


        /// <summary>
        ///     Determins in resource strings and images are cached or always retrieved.
        ///     The default is to not cache resources.
        /// </summary>
        public bool CacheResourecs
        {
            get => GetState(UserActionStates.CacheResources);
            set
            {
                SetState(UserActionStates.CacheResources, value, raisePropertyChange: false);
                SetState(UserActionStates.SmallImageResourceCached | UserActionStates.LargeImageResourceCached,
                         false, raisePropertyChange: false);
            }
        }


        /// <summary>
        ///     Returns all the currently stored property keys.
        /// </summary>
        public IEnumerable<string> Properties
        {
            get
            {
                IEnumerator keys;
                lock (_properties)
                    keys = _properties.GetEnumerator();

                while (keys.MoveNext())
                    yield return (string)keys.Current;
            }
        }


        /// <summary>
        ///     Returns or sets the image resource name to use if it is different to the UserAction.Id.
        /// </summary>
        public string AlternateImageResource
        {
            get => GetPropertyValue<string>();
            set
            {
                if (value != null &&
                    (value.EndsWith("32x32") || value.EndsWith("16x16") ||
                     value.EndsWith("48x48")))
                {
                    value = value.Substring(startIndex: 0, length: 5);
                }

                SetPropertyValue(value);
                RefreshImages();
            }
        }


        /// <summary>
        ///     The event handler called when a UserAction property changes.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged
        {
            add => _propertyChanged += value;
            // ReSharper disable once DelegateSubtraction
            remove => _propertyChanged -= value;
        }

        /// <summary>
        ///     Gets the string value for the named property, If the property value is null it will raise the UserActionGetString
        ///     event to retrieve the string value
        /// </summary>
        /// <param name="name"></param>

        protected string GetPropertyOrResourceString([CallerMemberName] string name = null)
        {
            // ReSharper disable once ExplicitCallerInfoArgument
            var value = GetPropertyValue<string>(name);
            if (value != null || !UserActionManager.UseResourceStrings)
                return value;

            var e = new UserActionGetResourceEventArgs(this, name);

            OnGetResource(e);
            var stringValue = e.Value as string;

            if (CacheResourecs)
                SetPropertyValue(stringValue, name);

            return stringValue;
        }

        /// <summary>
        ///     Raises the UserActionManager.UserActionGetResource event.
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnGetResource(UserActionGetResourceEventArgs e) => UserActionManager.OnUserActionGetResource(e);


        /// <summary>
        ///     The event handler called when the Execute method is called
        /// </summary>
        public event EventHandler<UserActionExecuteEventArgs> Executed
        {
            add => _executed += value;
            // ReSharper disable once DelegateSubtraction
            remove => _executed -= value;
        }


        internal static string GetName(string id)
        {
            if (string.IsNullOrEmpty(id))
                return null;

            var namePartIndex = id.LastIndexOf(value: '_');
            return namePartIndex == -1 ? id : id.Substring(namePartIndex + 1);
        }


        /// <summary>
        ///     Raises the Executed event
        /// </summary>
        /// <param name="e">The UserActionExecuteEventArgs to pass to the event.</param>
        protected virtual void OnExecuted(UserActionExecuteEventArgs e) => _executed.Invoke(this, e);


        /// <summary>
        ///     Raises the PropertyChanged event.
        /// </summary>
        /// <param name="propertyName">The propertyName of the property which has changed.</param>
        protected virtual void OnPropertyChanged(string propertyName)
            => _propertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));


        /// <summary>
        ///     Returns the value of a built-in property.
        /// </summary>
        /// <param name="name">The name of the property value to get.</param>
        /// <param name="value">Filled with value of the property on success.</param>
        [SuppressMessage("Microsoft.Design", "CA1007:UseGenericsWhereAppropriate")]
        [SuppressMessage("Microsoft.Design", "CA1045:DoNotPassTypesByReference", MessageId = "1#")]
        protected virtual bool GetBuiltInProperty(string name, out object value)
        {
            value = null;
            switch (name)
            {
                case nameof(Id):
                    value = Id;
                    break;
                case nameof(Text):
                    value = GetPropertyValue<object>(nameof(Text));
                    break;
                case RegularImage:
                    // ReSharper disable once ExplicitCallerInfoArgument
                    value = GetPropertyValue<object>(RegularImage);
                    break;
                case LargeImage:
                    // ReSharper disable once ExplicitCallerInfoArgument
                    value = GetPropertyValue<object>(LargeImage);
                    break;
                case nameof(ToolTip):
                    value = GetPropertyValue<object>(nameof(ToolTip));
                    break;
                case nameof(Enabled):
                    value = Enabled;
                    break;
                case nameof(StateProperty):
                    // ReSharper disable once ExplicitCallerInfoArgument
                    value = GetPropertyValue<UserActionStates>(StateProperty);
                    break;
                case nameof(Visible):
                    value = Visible;
                    break;
                case ParentVisible:
                    value = GetState(UserActionStates.ParentVisible);
                    break;
                case ParentEnabled:
                    value = GetState(UserActionStates.ParentEnabled);
                    break;
                case nameof(AutoCheck):
                    value = AutoCheck;
                    break;
                default:
                    return false;
            }

            return true;
        }


        /// <summary>
        ///     Sets the value of a known built in property.
        /// </summary>
        /// <param name="name">The name of the property to set. This will be one of the values.</param>
        /// <param name="value">The new value for the property</param>
        /// <returns>True if the value was set;false otherwise</returns>
        [SuppressMessage("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily")]
        protected virtual bool SetBuiltInProperty(string name, object value)
        {
            switch (name)
            {
                case nameof(Id):
                    throw new ArgumentException("Id property is read-only", name);
                case nameof(Text):
                    Text = Convert.ToString(value);
                    break;
                case RegularImage:
                    SetImage(UserActionImageSize.Regular, value);
                    break;
                case LargeImage:
                    SetImage(UserActionImageSize.Large, value);
                    break;
                case nameof(ToolTip):
                    ToolTip = Convert.ToString(value);
                    break;
                case StateProperty:
                    throw new ArgumentException("State property is read-only", name);
                case nameof(Enabled):
                    Enabled = Convert.ToBoolean(value);
                    break;
                case nameof(Visible):
                    Visible = Convert.ToBoolean(value);
                    break;
                case ParentVisible:
                    // ReSharper disable once ExplicitCallerInfoArgument
                    SetState(UserActionStates.ParentVisible, Convert.ToBoolean(value), name);
                    break;
                case ParentEnabled:
                    // ReSharper disable once ExplicitCallerInfoArgument
                    SetState(UserActionStates.ParentEnabled, Convert.ToBoolean(value), name);
                    break;
                case nameof(UpdateDelegate):
                    throw new ArgumentException("Property name is read-only", name);
                case nameof(AutoCheck):
                    AutoCheck = Convert.ToBoolean(value);

                    break;
                default:
                    return false;
            }

            return true;
        }


        /// <summary>
        ///     Returns if a property exists in the UserAction.
        /// </summary>
        /// <param name="name">The name of the property to check</param>
        /// <returns>true if the property exists;false otherwise</returns>
        [DebuggerStepThrough]
        public virtual bool HasProperty(string name)
        {
            if (_properties == null)
                return false;

            lock (_properties)
                return _properties.Contains(name);
        }


        /// <summary>
        ///     Returns the value of a property contained in the UserAction
        /// </summary>
        /// <typeparam name="T">The type of property to retrieve.</typeparam>
        /// <param name="name">The name of the property to retrieve the value for.</param>
        /// <returns>The value of the property if it exists; otherwise the default type value is returned.</returns>
        protected T GetPropertyValue<T>([CallerMemberName] string name = null)
        {
            Check.NotNull(name, nameof(name));
            if (_properties == null)
                return default(T);

            lock (_properties)
            {
                // ReSharper disable once AssignNullToNotNullAttribute
                if (_properties.Contains(name))
                    return (T)_properties[name];
            }

            return default(T);
        }


        /// <summary>
        ///     Commits the name/value key pair to storage.
        /// </summary>
        /// <param name="name">The name of the property to store</param>
        /// <param name="value">The value of the property to store.</param>
        /// <param name="raisePropertyChange">True to raise the PropertyChange for the item.</param>
        protected bool SetPropertyValue<T>(T value, [CallerMemberName] string name = null, bool raisePropertyChange = true)
        {
            Check.NotNull(name, nameof(name));

            if (value == null && _properties == null)
                return false;

            if (GetBuiltInProperty(name, out var curValue) == false)
            {
                // ReSharper disable once ExplicitCallerInfoArgument
                curValue = GetPropertyValue<T>(name);
            }

            var equals = Equals(curValue, value);
            if (equals)
                return false;

            SetPropertyValueCore(name, value);
            if (raisePropertyChange)
                OnPropertyChanged(name);

            return true;
        }


        /// <summary>
        ///     Commits the name/value key pair to storage.
        /// </summary>
        /// <param name="name">The name of the property to store</param>
        /// <param name="value">The value of the property to store.</param>
        /// <param name="nullIsValid">
        ///     true if a null value is a valid value; otherwise a false value will remove the item from the
        ///     backing store.
        /// </param>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        protected void SetPropertyValueCore(string name, object value, bool nullIsValid = false)
        {
            lock (_properties)
            {
                if (value == null && nullIsValid == false)
                {
                    _properties.Remove(name);
                    if (_properties.Count == 0)
                        _properties = null;
                    return;
                }

                _properties[name] = value;
            }
        }


        /// <summary>
        ///     Raises the PropertyChanged event passing RefreshImages so that the update delegate can provide a different image
        ///     for the UserAction if required.
        /// </summary>
        public void RefreshImages()
        {
            SetState(UserActionStates.LargeImageResourceCached | UserActionStates.SmallImageResourceCached, false);
            OnPropertyChanged(nameof(RefreshImages));
        }


        private bool GetState(UserActionStates item)
            // ReSharper disable once ExplicitCallerInfoArgument
            => (GetPropertyValue<UserActionStates>(StateProperty) & item) == item;

        private void SetState(UserActionStates newState, bool value, [CallerMemberName] string propertyName = null, bool raisePropertyChange = true)
        {
            var state = (UserActionStates)GetProperty(StateProperty);
            state = value ? state | newState : state & ~newState;
            // ReSharper disable once ExplicitCallerInfoArgument
            if (SetPropertyValue(state, StateProperty, raisePropertyChange: false) & raisePropertyChange)
                OnPropertyChanged(propertyName);
        }


        /// <summary>
        ///     Returns the Image associated with the UserAction .
        /// </summary>

        /// <returns>Image</returns>
        /// <remarks>Raises the PropertyChanged event when the value changes.</remarks>
        public virtual object GetImage(UserActionImageSize imageSize, int deviceDpi = 0)
        {
            var propertyName = imageSize == UserActionImageSize.Regular ? RegularImage : LargeImage;

            // ReSharper disable once ExplicitCallerInfoArgument
            var img = GetPropertyValue<object>(propertyName);
            if (img != null || HasProperty(propertyName))
                return img;

            var cachedFlag = imageSize == UserActionImageSize.Large ? UserActionStates.LargeImageResourceCached : UserActionStates.SmallImageResourceCached;

            if (GetState(cachedFlag))
                return null;

            var id = HasProperty(nameof(AlternateImageResource))
                         ? (string)GetProperty(nameof(AlternateImageResource))
                         : Id;

            var e = new UserActionGetResourceEventArgs(this, id, propertyName, deviceDpi);

            OnGetResource(e);

            if (CacheResourecs)
            {
                SetImage(imageSize, e.Value);
                SetState(cachedFlag, true);
            }

            return e.Value;
        }


        /// <summary>
        ///     Sets the Image associated with the UserAction .
        /// </summary>

        /// <remarks>Raises the PropertyChanged event when the value changes.</remarks>
        public virtual void SetImage(UserActionImageSize imageSize, object value)
            // ReSharper disable once ExplicitCallerInfoArgument
            => SetPropertyValue(value, imageSize == UserActionImageSize.Regular ? RegularImage : LargeImage);


        /// <summary>
        ///     Raises the Executed event.
        /// </summary>
        /// <param name="parameter"></param>
        public virtual object Execute(object parameter = null)
        {
            if (UserActionManager.UserActions.Contains(Id))
                return UserActionManager.ExecuteUserAction(Id, parameter);

            ExecuteCore(autoCheck: true, parameter: parameter);
            return parameter;
        }


        internal virtual bool ExecuteCore(bool autoCheck, object parameter = null)
        {
            if (Executing)
                return false;

            var e = new UserActionExecuteEventArgs(this, parameter);
            ExecuteCore(e, autoCheck);
            return e.Handled;
        }

        internal void ExecuteCore(UserActionExecuteEventArgs e, bool autoCheck)
        {
            if (Executing)
                return;

            // Trace.TraceInformation("UserAction hash/id value: " & Me.GetHashCode().ToString & "; " & Me.Id.ToString)
            SetState(UserActionStates.Executing, value: true);
            try
            {
                if (autoCheck && AutoCheck)
                    Checked = !Checked;

                OnExecuted(e);
            }
            finally
            {
                SetState(UserActionStates.Executing, value: false);
            }
        }


        /// <summary>
        ///     If an updateDelegate is supplied during the construction of this class then this method will invoke the delegate.
        ///     The delegate will then update all the states of this UserAction.
        /// </summary>
        public virtual void Update() => Update(propertyName: null);

        /// <summary>
        ///     If an updateDelegate is supplied during the construction of this class then this method will invoke the delegate.
        ///     The delegate will then update only the state of the property with the who's name matches the propertyName passed
        ///     in..
        /// </summary>
        public virtual void Update(string propertyName) => UpdateDelegate?.Invoke(this, propertyName, UserActionManager.ActiveTarget?.Context);


        /// <summary>
        ///     Returns the value of the property matching the name passed in.
        /// </summary>
        /// <param name="name">The name of the property to retrieve the value from.</param>
        public object GetProperty(string name)
        {
            if (GetBuiltInProperty(name, out var value) == false)
                // ReSharper disable once ExplicitCallerInfoArgument
                value = GetPropertyValue<object>(name);

            return value;
        }

        /// <summary>
        ///     Returns the value of the property matching the name passed in.
        /// </summary>
        /// <param name="name">The name of the property to retrieve the value from.</param>
        /// <param name="defaultValue">The default value to return if the property does not exist.</param>
        public TValue GetProperty<TValue>(string name, TValue defaultValue)
        {
            if (!HasProperty(name))
                return defaultValue;

            if (GetBuiltInProperty(name, out var value) == false)
                // ReSharper disable once ExplicitCallerInfoArgument
                value = GetPropertyValue<object>(name);

            var converter = TypeDescriptor.GetConverter(typeof(TValue));
            if (converter.CanConvertTo(typeof(TValue)))
                return (TValue)converter.ConvertFrom(value);
            return defaultValue;
        }

        /// <summary>
        ///     Sets the value of the property matching the name passed in.
        /// </summary>
        /// <param name="name">The name of the property to retrieve the value from.</param>
        /// <param name="value">The new value to set the property to.</param>
        public void SetProperty(string name, object value)
        {
            if (SetBuiltInProperty(name, value) == false)
                // ReSharper disable once ExplicitCallerInfoArgument
                SetPropertyValue(value, name);
        }


        [Flags]
        private enum UserActionStates
        {
            //None = 0x0,
            Visible = 0x1,
            Enabled = 0x2,
            Checked = 0x4,
            Indeterminate = 0x8,
            Executing = 0x10,
            ParentVisible = 0x20,
            ParentEnabled = 0x40,
            AutoCheck = 0x80,
            ImageSizeLarge = 0x100,
            CacheResources = 0x200,
            LargeImageResourceCached = 0x400,
            SmallImageResourceCached = 0x800
        }
    }
}