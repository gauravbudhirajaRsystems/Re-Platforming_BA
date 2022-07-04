// © Copyright 2018 Levit & James, Inc.

namespace LevitJames.Core.Diagnostics
{
    /// <summary>
    ///     Defines the built-in diagnostic options for use with the Diagnostics class.
    ///     The values are defined as strings constants so additional constants can be defined outside of this Assembly.
    /// </summary>
    public static class AppDiagnosticOptions
    {
        /// <summary>
        ///     No Diagnostic Options
        /// </summary>
        public const string None = "";

        /// <summary>
        ///     When this option is set no Windows Hooks are installed.
        /// </summary>
        /// <remarks>With this option set the functionality may differ from the expected behavior.</remarks>
        public const string NoHooks = "NoHooks";

        /// <summary>
        ///     When this option is set no Idle Handlers are installed.
        /// </summary>
        /// <remarks>With this option set the functionality may differ from the expected behavior.</remarks>
        public const string NoIdleHandlers = "NoIdleHandlers";

        /// <summary>When this option is set no painting locks are used. This may result in more on screen flickering.</summary>
        public const string SuppressScreenLocking = "SuppressScreenLocking";
    }
}