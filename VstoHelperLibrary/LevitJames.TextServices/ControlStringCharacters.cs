// © Copyright 2018 Levit & James, Inc.

namespace LevitJames.TextServices
{
    /// <summary>
    ///     Defines the set of ASCII control characters as string types
    /// </summary>
    public static class ControlStringCharacters
    {
        /// <summary>
        ///     A char value representing a null/zero, Ascii character (0, 0x0).
        /// </summary>

        public const string Null = "\x0000";

        /// <summary>
        ///     A char value representing the Ascii start of heading character (1, 0x1).
        /// </summary>

        public const string StartOfHeading = "\x0001";

        /// <summary>
        ///     A char value representing the Ascii start of text character (2, 0x2).
        /// </summary>

        public const string StartOfText = "\x0002";

        /// <summary>
        ///     A char value representing the Ascii end Of text character (3, 0x3).
        /// </summary>

        public const string EndOfText = "\x0003";

        /// <summary>
        ///     A char value representing the Ascii end Of transmission character (4, 0x4).
        /// </summary>

        public const string EndOfTransmission = "\x0004";

        /// <summary>
        ///     A char value representing the Ascii inquiry character (5, 0x5).
        /// </summary>

        public const string Enquiry = "\x0005";

        /// <summary>
        ///     A char value representing the Ascii acknowledge character (6, 0x6).
        /// </summary>

        public const string Acknowledge = "\x0006";

        /// <summary>
        ///     A char value representing the Ascii bell character (7, 0x7).
        /// </summary>

        public const string Bell = "\x0007";

        /// <summary>
        ///     A char value representing the Ascii backspace character (8, 0x8).
        /// </summary>

        public const string Backspace = "\x0008";

        /// <summary>
        ///     A char value representing the Ascii horizontal tab character (9, 0x9).
        /// </summary>

        public const string HorizontalTab = "\x0009";

        /// <summary>
        ///     A char value representing the Ascii line feed character (10, 0xA).
        /// </summary>

        public const string Linefeed = "\x000A";

        /// <summary>
        ///     A char value representing the Ascii vertical tab character (11, 0xB).
        /// </summary>

        public const string VerticalTab = "\x000B";

        /// <summary>
        ///     A char value representing the Ascii form feed character (12, 0xC).
        /// </summary>

        public const string FormFeed = "\x000C";

        /// <summary>
        ///     A char value representing the Ascii return character (13, 0xD)..
        /// </summary>

        public const string CharacterReturn = "\x000D";

        /// <summary>
        ///     A char value representing the Ascii shift in character (14, 0xE).
        /// </summary>

        public const string ShiftIn = "\x000E";

        /// <summary>
        ///     A char value representing the Ascii shift out character (15, 0xF).
        /// </summary>

        public const string ShiftOut = "\x000F";

        /// <summary>
        ///     A char value representing the Ascii data link character(16, 0x10).
        /// </summary>

        public const string DataLink = "\x0010";

        /// <summary>
        ///     A char value representing the Ascii X on character (17, 0x11).
        /// </summary>

        public const string XOn = "\x0011";

        /// <summary>
        ///     A char value representing the Ascii device control 2 character (18, 0x12).
        /// </summary>

        public const string DeviceControl2 = "\x0012";

        /// <summary>
        ///     A char value representing the Ascii x off character (19, 0x13).
        /// </summary>

        public const string XOff = "\x0013";

        /// <summary>
        ///     A char value representing the Ascii device control 4 character (20, 0x14).
        /// </summary>

        public const string DeviceControl4 = "\x0014";

        /// <summary>
        ///     A char value representing the Ascii negative acknowledge character (21, 0x15)..
        /// </summary>

        public const string NegativeAcknowledge = "\x0015";

        /// <summary>
        ///     A char value representing the Ascii synchronous idle character (22, 0x16).
        /// </summary>

        public const string SynchronousIdle = "\x0016";

        /// <summary>
        ///     A char value representing the Ascii end of Transmission block character (23, 0x17).
        /// </summary>

        public const string EndTransmissionBlock = "\x0017";

        /// <summary>
        ///     A char value representing the Ascii cancel line character (24, 0x18).
        /// </summary>

        public const string CancelLine = "\x0018";

        /// <summary>
        ///     A char value representing the Ascii end of medium character (25, 0x19).
        /// </summary>

        public const string EndOfMedium = "\x0019";

        /// <summary>
        ///     A char value representing the Ascii substitute character (26, 0x1A).
        /// </summary>

        public const string Substitute = "\x001A";

        /// <summary>
        ///     A char value representing the Ascii escape character (27, 0x1B).
        /// </summary>

        public const string Escape = "\x001B";

        /// <summary>
        ///     A char value representing the Ascii file separator character (28, 0x1C).
        /// </summary>

        public const string FileSeparator = "\x001C";

        /// <summary>
        ///     A char value representing the Ascii group separator character (29, 0x1D).
        /// </summary>

        public const string GroupSeparator = "\x001D";

        /// <summary>
        ///     A char value representing the Ascii record separator character (30, 0x1E).
        /// </summary>

        public const string RecordSeparator = "\x001E";

        /// <summary>
        ///     A char value representing the Ascii unit separator character (31, 0x1F).
        /// </summary>

        public const string UnitSeparator = "\x001F";
    }
}