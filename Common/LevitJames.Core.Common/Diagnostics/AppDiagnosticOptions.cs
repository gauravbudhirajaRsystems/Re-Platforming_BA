namespace LevitJames.Core.Common.Diagnostics
{
    /// <summary>
    ///     Defines the built-in diagnostic options for use with the Diagnostics class.
    ///     The values are defined as strings constants so additional constants can be defined outside of this Assembly.
    /// </summary>
    public static class AppDiagnosticOptions
    {
        /// <summary>When this option is set no painting locks are used. This may result in more on screen flickering.</summary>
        public const string SuppressScreenLocking = "SuppressScreenLocking";
    }
}
