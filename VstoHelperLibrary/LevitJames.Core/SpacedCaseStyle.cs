// © Copyright 2018 Levit & James, Inc.

using System;

namespace LevitJames.Core
{
    /// <summary>
    ///     Defines how the ToSpacedCase extension method changes the casing of its passed string.
    /// </summary>
    [Flags]
    public enum SpacedCaseStyle
    {
        /// <summary>
        ///     The formatting is unchanged.
        /// </summary>
        None = 0,

        /// <summary>
        ///     Only the first letter of the first word is capitalized.  For example "the quick brown fox" is formatted as "The
        ///     quick brown fox"
        /// </summary>
        CapitalizeFirstWordOnly = 1,

        /// <summary>
        ///     Each first letter of each word is capitalized. For example "the quick brown fox" is formatted as "The Quick Brown
        ///     Fox"
        /// </summary>
        CapitalizeAllWords = 2
    }
}