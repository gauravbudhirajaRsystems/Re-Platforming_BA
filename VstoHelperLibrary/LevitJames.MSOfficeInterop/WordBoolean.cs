// © Copyright 2018 Levit & James, Inc.

namespace LevitJames.MSOffice.MSWord
{
    /// <summary>
    ///     Defines the values of a Word property that returns a boolean. On occasions Word may not be able to return a valid
    ///     value, and instead will throw an exception.
    ///     Instead of throwing an exception many of the Word extension methods return a WordBoolean and set the value to
    ///     WordBoolean.Unknown.
    /// </summary>
    public enum WordBoolean
    {
        /// <summary>
        ///     The property state is False
        /// </summary>

        False = 0,

        /// <summary>
        ///     The property is True
        /// </summary>

        True = -1,

        /// <summary>
        ///     The property state cannot be determined at this time.
        /// </summary>
        /// <remarks>
        ///     This value can be returned if Word is in a state where the call would normally throw an exception. For example
        ///     if Word has a floating modaless dialog open, like the Find and Replace dialog, many Word properties throw
        ///     exceptions.
        /// </remarks>
        Unknown = 2
    }
}
