// © Copyright 2018 Levit & James, Inc.

using System;
using JetBrains.Annotations;

namespace LevitJames.Core
{
    /// <summary>
    /// </summary>
    public class UserActionGetResourceEventArgs : EventArgs
    {
        private readonly string _id;


        /// <summary>
        ///     Creates a new instance of the UserActionLoadImageEventArgs using the supplied members.
        /// </summary>
        /// <param name="userAction">The userAction instance to load the image for.</param>
        /// <param name="propertyName">The name of the string resource to load.</param>
        /// <param name="deviceDpi"></param>
        public UserActionGetResourceEventArgs([NotNull] UserAction userAction, string propertyName, int deviceDpi) : this(userAction, userAction.Id, propertyName, deviceDpi) { }

        /// <summary>
        ///     Creates a new instance of the UserActionLoadImageEventArgs using the supplied members.
        /// </summary>
        /// <param name="userAction">The userAction instance to load the image for.</param>
        /// <param name="propertyName">The name of the string resource to load.</param>
        public UserActionGetResourceEventArgs([NotNull] UserAction userAction, string propertyName) : this(userAction, userAction.Id, propertyName, 0) { }

        /// <summary>
        ///     Creates a new instance of the UserActionLoadImageEventArgs using the supplied members.
        /// </summary>
        /// <param name="userAction">The userAction instance to load the image for.</param>
        /// <param name="resourceId">The resourceId to retrieve.</param>
        /// <param name="propertyName">The name of the string resource to load.</param>
        /// <param name="deviceDpi"></param>
        public UserActionGetResourceEventArgs([NotNull] UserAction userAction, string resourceId, string propertyName, int deviceDpi)
        {
            Check.NotNull(userAction, nameof(userAction));
            UserAction = userAction;
            _id = resourceId;
            DeviceDpi = deviceDpi;
            PropertyName = propertyName;
        }

        /// <summary>
        ///     Creates a new instance of the UserActionLoadImageEventArgs using the supplied members.
        /// </summary>
        /// <param name="userActionId">The id of the UserAction to load the image for.</param>
        /// <param name="propertyName">The name of the string resource to load.</param>
        public UserActionGetResourceEventArgs([NotNull] string userActionId, string propertyName)
        {
            Check.NotNull(userActionId, nameof(userActionId));
            _id = userActionId;
            PropertyName = propertyName;
        }

        /// <summary>
        ///     Returns the Id from the UserAction, or a string Id if the UserAction is a simple userAction (A userAction
        ///     represented by a string id only)
        /// </summary>
        public string Id => _id ?? UserAction.Id;

        /// <summary>
        ///     Returns the UserAction to load the Image for.
        /// </summary>
        public UserAction UserAction { get; }

        /// <summary>
        ///     The property used to store the Image to be loaded for the UserAction.
        /// </summary>
        public object Value { get; set; }

        /// <summary>
        ///     The size of the image requested to load.
        /// </summary>
        public string PropertyName { get; }

        
        /// <summary>
        ///     The dpi of the image requested.
        /// </summary>
        public int DeviceDpi { get; }
    }
}