// © Copyright 2018 Levit & James, Inc.

using System;

namespace LevitJames.Core.Diagnostics
{
    /// <summary>
    ///     The event arguments that are used in the AppDiagnostics.OptionChanged event.
    /// </summary>
    public class AppDiagnosticOptionsChangedEventArgs : EventArgs
    {
        /// <summary>
        ///     Creates a new instance of AppDiagnosticOptionsChangedEventArgs
        /// </summary>
        /// <param name="itemChanged"></param>
        public AppDiagnosticOptionsChangedEventArgs(string itemChanged)
        {
            ItemChanged = itemChanged;
        }

        /// <summary>
        ///     The name of the option that changed.
        /// </summary>
        public string ItemChanged { get; }

        /// <summary>A helper method to check if the option name passed matches the name in the ItemChanged property.</summary>
        /// <param name="name">The name of the option to compare.</param>
        /// <returns>True if the name argument matches the ItemChanged value.</returns>
        public bool HasItemChanged(string name)
        {
            return string.IsNullOrEmpty(ItemChanged) ||
                   string.Equals(name, ItemChanged, StringComparison.OrdinalIgnoreCase);
        }
    }
}