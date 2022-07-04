// © Copyright 2018 Levit & James, Inc.

namespace LevitJames.Core
{
    /// <summary>
    ///     Defines the members that represent the checked state of a UserAction.
    /// </summary>
    public enum UserActionCheckState
    {
        /// <summary>
        ///     The control is unchecked.
        /// </summary>
        Unchecked = 0,

        /// <summary>
        ///     The control is checked.
        /// </summary>
        Checked = 1,

        /// <summary>
        ///     The control is indeterminate. An indeterminate control UI element typically has a shaded appearance.
        /// </summary>
        Indeterminate = 2
    }
}