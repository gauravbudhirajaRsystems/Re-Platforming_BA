// © Copyright 2018 Levit & James, Inc.

using System;

namespace LevitJames.Core
{
    /// <summary>Enum for tracking emphasis character formatting for saving to ranges in a document.</summary>
    [Flags]
#pragma warning disable CA1714 // Flags enums should have plural names
    public enum CharacterFormattingEmphasis
#pragma warning restore CA1714 // Flags enums should have plural names
    {
        /// <summary>If any attribute found, returns all as true</summary>
        MatchAnyBit = unchecked((int)0x80000000),

        /// <summary></summary>
        Unknown = unchecked((int)0xFFFFFFFF),

        /// <summary></summary>
        AllBits1 = unchecked((int)0xFFFFFFFF),

        /// <summary>The underline bit. Single underlining only.</summary>
        Underline = 0x2,

        /// <summary>The italic bit.</summary>
        Italic = 0x4,

        /// <summary>The small caps bit.</summary>
        SmallCaps = 0x8,

        /// <summary></summary>
        // ReSharper disable once InconsistentNaming
        // ??? Rename to UnderlineAndItalic?
        UI = 0x6,

        /// <summary></summary>
        // ReSharper disable once InconsistentNaming
        // ??? Rename
        UIS = 0xE,

        /// <summary></summary>
        Any = 0xF,

        /// <summary>Returns all as false</summary>
        None = 0
    }
}