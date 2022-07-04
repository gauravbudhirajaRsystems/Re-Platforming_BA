// © Copyright 2018 Levit & James, Inc.

namespace LevitJames.Core
{
    /// <summary>
    ///     Stores the difference in stored values when comparing two serialized streams.
    /// </summary>
    public struct SerizaliationStreamComparerItem
    {
        /// <summary>
        ///     The serialized key used to store the value
        /// </summary>
        public string Key { get; internal set; }

        /// <summary>
        ///     The value from the first serialized stream
        /// </summary>
        public object Value1 { get; internal set; }

        /// <summary>
        ///     The value from the second serialized stream
        /// </summary>
        public object Value2 { get; internal set; }
    }
}