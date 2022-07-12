using LevitJames.Shared.Common;
using System.Collections.Generic;

namespace LevitJames.Core.Common.Diagnostics
{
    /// <summary>
    ///     A class for managing diagnostic options
    /// </summary>
    public static class AppDiagnostics
    {
        private static readonly Dictionary<string, bool?> Options = new Dictionary<string, bool?>();


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
    }
}
