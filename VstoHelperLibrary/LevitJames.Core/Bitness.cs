// © Copyright 2018 Levit & James, Inc.

namespace LevitJames.Core
{
	/// <summary>
    ///     Values that represent the bitness of an machine file.
    /// </summary>
    public enum Bitness
    {
        /// <summary>The bitness is unknown</summary>
        Unknown,

        /// <summary>
        ///     The bitness will be determined at runtime.
        /// </summary>
        AnyCPU,

        /// <summary>
        ///     The bitness is 32bit
        /// </summary>
        // ReSharper disable once InconsistentNaming
        x86,

        /// <summary>
        ///     The bitness is 64bit
        /// </summary>
        // ReSharper disable once InconsistentNaming
        x64
    }
}