// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using JetBrains.Annotations;

namespace LevitJames.Core.Diagnostics
{
    /// <summary>
    ///     A class for managing diagnostic options
    /// </summary>
    public static class AppDiagnostics
    {
        private static readonly Dictionary<string, bool?> Options = new Dictionary<string, bool?>();

        /// <summary>
        ///     An Event that is raised when an option is added/removed or changed.
        /// </summary>
        public static event EventHandler<AppDiagnosticOptionsChangedEventArgs> OptionChanged;


        /// <summary>
        ///     Returns the value of the option with the provided name.
        ///     If the option does not exist null is returned.
        /// </summary>
        /// <param name="name">The name of the option to retrieve</param>
        /// <returns>The value for the option, or null if the option does not exist.</returns>
        public static bool GetOption([NotNull] string name)
        {
            return Options.TryGetValue(name, out bool? value) && value.GetValueOrDefault();
        }

        /// <summary>
        ///     Returns the value of the option with the provided name.
        ///     If the option does not exist null is returned.
        /// </summary>
        /// <param name="name">The name of the option to retrieve</param>
        /// <returns>The value for the option, or null if the option does not exist.</returns>
        public static bool? GetOptionOrNull([NotNull] string name)
        {
            return Options.TryGetValue(name, out bool? value) ? value : null;
        }

        /// <summary>
        ///     Sets the value of an option using the provided name and value.
        ///     If the value is null the option is removed.
        /// </summary>
        /// <param name="name">The name of the option to set.</param>
        /// <param name="value">The value for the option. If the value is null the option is removed.</param>
        public static void SetOption([NotNull] string name, bool? value)
        {
            Check.NotEmpty(name, "name");

            bool changed;
            if (value == null)
            {
                changed = Options.Remove(name);
            }
            else
            {
                if (Options.TryGetValue(name, out var existingValue))
                {
                    changed = !value.Equals(existingValue);
                    if (changed)
                    {
                        Options[name] = value;
                    }
                }
                else
                {
                    if (Regex.IsMatch(name, "^[a-zA-Z][a-zA-Z0-9]*$") == false)
                    {
                        throw new ArgumentException(
                                                    "The name provided is invalid. Name can only contain alphanumeric characters", nameof(name));
                    }

                    Options.Add(name, value);
                    changed = true;
                }
            }

            if (changed)
            {
                OnOptionChanged(name);
            }
        }


        /// <summary>
        ///     Returns if the option with the supplied name is available.
        /// </summary>
        /// <param name="name">The name of the option to check.</param>
        public static bool Contains(string name)
        {
            return Options.ContainsKey(name);
        }


        /// <summary>
        ///     Clears all the option values.;
        /// </summary>
        public static void Clear()
        {
            if (Options.Count <= 0)
                return;

            Options.Clear();
            OnOptionChanged(name: null);
        }


        private static void OnOptionChanged(string name)
        {
            OptionChanged?.Invoke(sender: null, e: new AppDiagnosticOptionsChangedEventArgs(name));
        }


        /// <summary>
        ///     Returns all the options defined in AppDiagnostics.
        /// </summary>
        public new static string ToString()
        {
            var sb = new StringBuilder();
            foreach (var kvp in Options)
            {
                sb.Append(kvp.Key);
                sb.Append(value: ':');
                sb.Append(kvp.Value);
                sb.Append(",");
            }

            if (sb.Length > 0)
            {
                sb.Remove(sb.Length - 1, length: 1);
            }

            return sb.ToString();
        }


        /// <summary>
        ///     Adds all the options in the provided comma delimited string in the format [Name]:[BooleanValue]
        /// </summary>
        /// <param name="options">A comma delimited string of name value pairs</param>
        /// <param name="append">append to add the options to the existing Options; false to clear the existing values.</param>
        /// <remarks>
        ///     The [BooleanValue] is optional if it is not provided then the value assumed is true. Example
        ///     'NoHooks:true,NoIdleHandler,SupressScreenUpdating:false'
        /// </remarks>
        public static void FromString(string options, bool append)
        {
            if (!string.IsNullOrEmpty(options))
            {
                options = options.Trim();
            }

            if (string.IsNullOrEmpty(options))
            {
                Clear();
                return;
            }

            if (append == false)
            {
                Clear();
            }

            var optionNameValues = options.Split(new[] {','}, StringSplitOptions.RemoveEmptyEntries);

            foreach (var nv in optionNameValues)
            {
                var nameValue = nv.Split(':');
                if (nameValue.Length > 0)
                {
                    nameValue[0] = nameValue[0].Trim();
                }

                if (nameValue.Length == 0 || string.IsNullOrEmpty(nameValue[0]))
                {
                    throw new ArgumentException("Invalid options string passed.", nameof(options));
                }

                var value = nameValue.Length == 1 || string.IsNullOrEmpty(nameValue[1]);
                if (value == false)
                {
                    if (bool.TryParse(nameValue[1], out value) == false)
                    {
                        throw new ArgumentException("Invalid options string passed.", nameof(options));
                    }
                }

                SetOption(nameValue[0], value);
            }
        }
    }
}