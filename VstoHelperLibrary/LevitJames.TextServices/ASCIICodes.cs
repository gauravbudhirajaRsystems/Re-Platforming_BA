// © Copyright 2018 Levit & James, Inc.

using System;
using System.Diagnostics.CodeAnalysis;

namespace LevitJames.TextServices
{
    /// <summary>
    ///     A class the represents all the character codes in the ASCII character set
    /// </summary>

    public static class AsciiCodes
    {
        // ControlCharacters
        /// <summary>
        ///     A byte value representing the first available Ascii character.
        /// </summary>

        public const byte MinValue = 0;

        /// <summary>
        ///     A byte value representing the last available Ascii character.
        /// </summary>

        public const byte MaxValue = Delete; //127

        /// <summary>
        ///     A byte value representing the last available Ascii control character.
        /// </summary>

        public const byte ControlCharactersStart = 0;

        /// <summary>
        ///     A byte value representing a null/zero, Ascii character.
        /// </summary>

        public const byte Null = 0;

        /// <summary>
        ///     A byte value representing the Ascii start of heading character.
        /// </summary>

        public const byte StartOfHeading = 1;

        /// <summary>
        ///     A byte value representing the Ascii start of text character.
        /// </summary>

        public const byte StartOfText = 2;

        /// <summary>
        ///     A byte value representing the Ascii end Of text character.
        /// </summary>

        public const byte EndOfText = 3;

        /// <summary>
        ///     A byte value representing the Ascii end Of transmission character.
        /// </summary>

        public const byte EndOfTransmission = 4;

        /// <summary>
        ///     A byte value representing the Ascii inquiry character.
        /// </summary>

        public const byte Enquiry = 5;

        /// <summary>
        ///     A byte value representing the Ascii acknowledge character.'
        /// </summary>

        public const byte Acknowledge = 6;

        /// <summary>
        ///     A byte value representing the Ascii bell character.
        /// </summary>

        public const byte Bell = 7;

        /// <summary>
        ///     A byte value representing the Ascii backspace character.
        /// </summary>

        public const byte Backspace = 8;

        /// <summary>
        ///     A byte value representing the Ascii horizontal tab character.
        /// </summary>

        public const byte HorizontalTab = 9;

        /// <summary>
        ///     A byte value representing the Ascii line feed character.
        /// </summary>

        public const byte Linefeed = 10;

        /// <summary>
        ///     A byte value representing the Ascii vertical tab character.
        /// </summary>

        public const byte VerticalTab = 11;

        /// <summary>
        ///     A byte value representing the Ascii form feed character.
        /// </summary>

        public const byte FormFeed = 12;

        /// <summary>
        ///     A byte value representing the Ascii return character.
        /// </summary>

        public const byte CharacterReturn = 13;

        /// <summary>
        ///     A byte value representing the Ascii shift in character.
        /// </summary>

        public const byte ShiftIn = 14;

        /// <summary>
        ///     A byte value representing the Ascii shift out character.
        /// </summary>

        public const byte ShiftOut = 15;

        /// <summary>
        ///     A byte value representing the Ascii data link character.
        /// </summary>

        public const byte DataLink = 16;

        /// <summary>
        ///     A byte value representing the Ascii X on character.
        /// </summary>

        public const byte XOn = 17;

        /// <summary>
        ///     A byte value representing the Ascii device control 2 character.
        /// </summary>

        public const byte DeviceControl2 = 18;

        /// <summary>
        ///     A byte value representing the Ascii x off character.
        /// </summary>

        public const byte XOff = 19;

        /// <summary>
        ///     A byte value representing the Ascii device control 4 character.
        /// </summary>

        public const byte DeviceControl4 = 20;

        /// <summary>
        ///     A byte value representing the Ascii negative acknowledge character.
        /// </summary>

        public const byte NegativeAcknowledge = 21;

        /// <summary>
        ///     A byte value representing the Ascii synchronous idle character.
        /// </summary>

        public const byte SynchronousIdle = 22;

        /// <summary>
        ///     A byte value representing the Ascii end of Transmission block character.
        /// </summary>

        public const byte EndTransmissionBlock = 23;

        /// <summary>
        ///     A byte value representing the Ascii cancel line character.
        /// </summary>

        public const byte CancelLine = 24;

        /// <summary>
        ///     A byte value representing the Ascii end of medium character.
        /// </summary>

        public const byte EndOfMedium = 25;

        /// <summary>
        ///     A byte value representing the Ascii substitute character.
        /// </summary>

        public const byte Substitute = 26;

        /// <summary>
        ///     A byte value representing the Ascii escape character.
        /// </summary>

        public const byte Escape = 27;

        /// <summary>
        ///     A byte value representing the Ascii file separator character.
        /// </summary>

        public const byte FileSeparator = 28;

        /// <summary>
        ///     A byte value representing the Ascii group separator character.
        /// </summary>

        public const byte GroupSeparator = 29;

        /// <summary>
        ///     A byte value representing the Ascii record separator character.
        /// </summary>

        public const byte RecordSeparator = 30;

        /// <summary>
        ///     A byte value representing the Ascii unit separator character.
        /// </summary>

        public const byte UnitSeparator = 31;

        /// <summary>
        ///     A byte value representing the Ascii control characters end character
        /// </summary>

        public const byte ControlCharactersEnd = UnitSeparator;

        // Control Characters End

        /// <summary>
        ///     A byte value representing the Ascii characters ' '
        /// </summary>

        public const byte Space = 32;

        /// <summary>
        ///     A byte value representing the Ascii character '!'
        /// </summary>

        public const byte ExclamationMark = 33;

        /// <summary>
        ///     A byte value representing the Ascii character '"'
        /// </summary>

        public const byte QuotationMark = 34;

        /// <summary>
        ///     A byte value representing the Ascii character '$'
        /// </summary>

        public const byte DollarSign = 36;

        /// <summary>
        ///     A byte value representing the Ascii character '%'
        /// </summary>

        public const byte PercentSign = 37;

        /// <summary>
        ///     A byte value representing the Ascii character '&amp;'
        /// </summary>

        public const byte Ampersand = 38;

        /// <summary>
        ///     A byte value representing the Ascii character '''
        /// </summary>

        public const byte Apostrophe = 39;

        /// <summary>
        ///     A byte value representing the Ascii character '('
        /// </summary>

        public const byte OpeningParentheses = 40;

        /// <summary>
        ///     A byte value representing the Ascii character ')'
        /// </summary>

        public const byte ClosingParentheses = 41;

        /// <summary>
        ///     A byte value representing the Ascii character '*'
        /// </summary>

        public const byte Asterisk = 42;

        /// <summary>
        ///     A byte value representing the Ascii character '+'
        /// </summary>

        public const byte Plus = 43;

        /// <summary>
        ///     A byte value representing the Ascii character ','
        /// </summary>

        public const byte Comma = 44;

        /// <summary>
        ///     A byte value representing the Ascii character '-'
        /// </summary>

        public const byte Hyphen = 45;

        /// <summary>
        ///     A byte value representing the Ascii character '.'
        /// </summary>

        public const byte Period = 46;

        /// <summary>
        ///     A byte value representing the Ascii character '/'
        /// </summary>

        public const byte ForwardSlash = 47;

        /// <summary>
        ///     A byte value representing the Ascii character '#'
        /// </summary>

        public const byte Crosshatch = 35;

        /// <summary>
        ///     A byte value representing the Ascii character '0'
        /// </summary>

        public const byte Zero = 48;

        /// <summary>
        ///     A byte value representing the Ascii character '1'
        /// </summary>

        public const byte One = 49;

        /// <summary>
        ///     A byte value representing the Ascii character '2'
        /// </summary>

        public const byte Two = 50;

        /// <summary>
        ///     A byte value representing the Ascii character ''
        /// </summary>

        public const byte Three = 51;

        /// <summary>
        ///     A byte value representing the Ascii character '4'
        /// </summary>

        public const byte Four = 52;

        /// <summary>
        ///     A byte value representing the Ascii character '5'
        /// </summary>

        public const byte Five = 53;

        /// <summary>
        ///     A byte value representing the Ascii character '6'
        /// </summary>

        public const byte Six = 54;

        /// <summary>
        ///     A byte value representing the Ascii character '7'
        /// </summary>

        public const byte Seven = 55;

        /// <summary>
        ///     A byte value representing the Ascii character '8'
        /// </summary>

        public const byte Eight = 56;

        /// <summary>
        ///     A byte value representing the Ascii character '9'
        /// </summary>

        public const byte Nine = 57;

        /// <summary>
        ///     A byte value representing the Ascii character ':'
        /// </summary>

        public const byte Colon = 58;

        /// <summary>
        ///     A byte value representing the Ascii character ';'
        /// </summary>

        public const byte Semicolon = 59;

        /// <summary>
        ///     A byte value representing the Ascii character '&lt;'
        /// </summary>

        public const byte LessThanSign = 60;

        /// <summary>
        ///     A byte value representing the Ascii character '='
        /// </summary>

        public const byte EqualsSign = 61;

        /// <summary>
        ///     A byte value representing the Ascii character '&lt;'
        /// </summary>

        public const byte GreaterThanSign = 62;

        /// <summary>
        ///     A byte value representing the Ascii character '?'
        /// </summary>

        public const byte QuestionMark = 63;

        /// <summary>
        ///     A byte value representing the Ascii character '@'
        /// </summary>

        public const byte AtSign = 64;

        /// <summary>
        ///     A byte value representing the Ascii character 'A'
        /// </summary>

        public const byte UppercaseA = 65;

        /// <summary>
        ///     A byte value representing the Ascii character 'B'
        /// </summary>

        public const byte UppercaseB = 66;

        /// <summary>
        ///     A byte value representing the Ascii character 'C'
        /// </summary>

        public const byte UppercaseC = 67;

        /// <summary>
        ///     A byte value representing the Ascii character 'D'
        /// </summary>

        public const byte UppercaseD = 68;

        /// <summary>
        ///     A byte value representing the Ascii character 'E'
        /// </summary>

        public const byte UppercaseE = 69;

        /// <summary>
        ///     A byte value representing the Ascii character 'F'
        /// </summary>

        public const byte UppercaseF = 70;

        /// <summary>
        ///     A byte value representing the Ascii character 'G'
        /// </summary>

        public const byte UppercaseG = 71;

        /// <summary>
        ///     A byte value representing the Ascii character 'H'
        /// </summary>

        public const byte UppercaseH = 72;

        /// <summary>
        ///     A byte value representing the Ascii character 'I'
        /// </summary>

        public const byte UppercaseI = 73;

        /// <summary>
        ///     A byte value representing the Ascii character 'J'
        /// </summary>

        public const byte UppercaseJ = 74;

        /// <summary>
        ///     A byte value representing the Ascii character 'K'
        /// </summary>

        public const byte UppercaseK = 75;

        /// <summary>
        ///     A byte value representing the Ascii character 'L'
        /// </summary>

        public const byte UppercaseL = 76;

        /// <summary>
        ///     A byte value representing the Ascii character 'M'
        /// </summary>

        public const byte UppercaseM = 77;

        /// <summary>
        ///     A byte value representing the Ascii character 'N'
        /// </summary>

        public const byte UppercaseN = 78;

        /// <summary>
        ///     A byte value representing the Ascii character 'O'
        /// </summary>

        public const byte UppercaseO = 79;

        /// <summary>
        ///     A byte value representing the Ascii character 'P'
        /// </summary>

        public const byte UppercaseP = 80;

        /// <summary>
        ///     A byte value representing the Ascii character 'Q'
        /// </summary>

        public const byte UppercaseQ = 81;

        /// <summary>
        ///     A byte value representing the Ascii character 'R'
        /// </summary>

        public const byte UppercaseR = 82;

        /// <summary>
        ///     A byte value representing the Ascii character 'S'
        /// </summary>

        public const byte UppercaseS = 83;

        /// <summary>
        ///     A byte value representing the Ascii character 'T'
        /// </summary>

        public const byte UppercaseT = 84;

        /// <summary>
        ///     A byte value representing the Ascii character 'U'
        /// </summary>

        public const byte UppercaseU = 85;

        /// <summary>
        ///     A byte value representing the Ascii character 'V'
        /// </summary>

        public const byte UppercaseV = 86;

        /// <summary>
        ///     A byte value representing the Ascii character 'W'
        /// </summary>

        public const byte UppercaseW = 87;

        /// <summary>
        ///     A byte value representing the Ascii character 'X'
        /// </summary>

        public const byte UppercaseX = 88;

        /// <summary>
        ///     A byte value representing the Ascii character 'Y'
        /// </summary>

        public const byte UppercaseY = 89;

        /// <summary>
        ///     A byte value representing the Ascii character 'Z'
        /// </summary>

        public const byte UppercaseZ = 90;


        /// <summary>
        ///     A byte value representing the Ascii character '['
        /// </summary>

        public const byte OpeningSquareBracket = 91;

        /// <summary>
        ///     A byte value representing the Ascii character '\'
        /// </summary>

        public const byte Backslash = 92;

        /// <summary>
        ///     A byte value representing the Ascii character ']'
        /// </summary>

        public const byte ClosingSquareBracket = 93;

        /// <summary>
        ///     A byte value representing the Ascii character '^'
        /// </summary>

        public const byte Caret = 94;

        /// <summary>
        ///     A byte value representing the Ascii character '_'
        /// </summary>

        public const byte Underscore = 95;

        /// <summary>
        ///     A byte value representing the Ascii character '''
        /// </summary>

        public const byte SingleQuote = 96;

        /// <summary>
        ///     A byte value representing the Ascii character 'a'
        /// </summary>

        public const byte LowercaseA = 97;

        /// <summary>
        ///     A byte value representing the Ascii character 'b'
        /// </summary>

        public const byte LowercaseB = 98;

        /// <summary>
        ///     A byte value representing the Ascii character 'c'
        /// </summary>

        public const byte LowercaseC = 99;

        /// <summary>
        ///     A byte value representing the Ascii character 'd'
        /// </summary>

        public const byte LowercaseD = 100;

        /// <summary>
        ///     A byte value representing the Ascii character 'e'
        /// </summary>

        public const byte LowercaseE = 101;

        /// <summary>
        ///     A byte value representing the Ascii character 'f'
        /// </summary>

        public const byte LowercaseF = 102;

        /// <summary>
        ///     A byte value representing the Ascii character 'gh'
        /// </summary>

        public const byte LowercaseG = 103;

        /// <summary>
        ///     A byte value representing the Ascii character 'h'
        /// </summary>

        public const byte LowercaseH = 104;

        /// <summary>
        ///     A byte value representing the Ascii character 'i'
        /// </summary>

        public const byte LowercaseI = 105;

        /// <summary>
        ///     A byte value representing the Ascii character 'j'
        /// </summary>

        public const byte LowercaseJ = 106;

        /// <summary>
        ///     A byte value representing the Ascii character 'k'
        /// </summary>

        public const byte LowercaseK = 107;

        /// <summary>
        ///     A byte value representing the Ascii character 'l'
        /// </summary>

        public const byte LowercaseL = 108;

        /// <summary>
        ///     A byte value representing the Ascii character 'm'
        /// </summary>

        public const byte LowercaseM = 109;

        /// <summary>
        ///     A byte value representing the Ascii character 'n'
        /// </summary>

        public const byte LowercaseN = 110;

        /// <summary>
        ///     A byte value representing the Ascii character 'o'
        /// </summary>

        public const byte LowercaseO = 111;

        /// <summary>
        ///     A byte value representing the Ascii character 'p'
        /// </summary>

        public const byte LowercaseP = 112;

        /// <summary>
        ///     A byte value representing the Ascii character 'q'
        /// </summary>

        public const byte LowercaseQ = 113;

        /// <summary>
        ///     A byte value representing the Ascii character 'r'
        /// </summary>

        public const byte LowercaseR = 114;

        /// <summary>
        ///     A byte value representing the Ascii character 's'
        /// </summary>

        public const byte LowercaseS = 115;

        /// <summary>
        ///     A byte value representing the Ascii character 't'
        /// </summary>

        public const byte LowercaseT = 116;

        /// <summary>
        ///     A byte value representing the Ascii character 'u'
        /// </summary>

        public const byte LowercaseU = 117;

        /// <summary>
        ///     A byte value representing the Ascii character 'v'
        /// </summary>

        public const byte LowercaseV = 118;

        /// <summary>
        ///     A byte value representing the Ascii character 'w'
        /// </summary>

        public const byte LowercaseW = 119;

        /// <summary>
        ///     A byte value representing the Ascii character 'x'
        /// </summary>

        public const byte LowercaseX = 120;

        /// <summary>
        ///     A byte value representing the Ascii character 'y'
        /// </summary>

        public const byte LowercaseY = 121;

        /// <summary>
        ///     A byte value representing the Ascii character 'z'
        /// </summary>

        public const byte LowercaseZ = 122;

        /// <summary>
        ///     A byte value representing the Ascii character '{'
        /// </summary>

        public const byte OpeningCurlyBrace = 123;

        /// <summary>
        ///     A byte value representing the Ascii character '|'
        /// </summary>

        public const byte VerticalLine = 124;

        /// <summary>
        ///     A byte value representing the Ascii character '}'
        /// </summary>

        public const byte ClosingCurlyBrace = 125;

        /// <summary>
        ///     A byte value representing the Ascii character '~'
        /// </summary>

        public const byte Tilde = 126;

        /// <summary>
        ///     A byte value representing the Ascii character ''
        /// </summary>

        public const byte Delete = 127;


        private const byte HexValue = 10;
        private const byte UpperLowerLetterOffet = 32;


        /// <summary>
        ///     Returns true if the character is a letter character
        /// </summary>
        /// <param name="value">The ascii character value to check.</param>
        /// <returns>true if the ascii character is a numeric character, false otherwise.</returns>

        public static bool IsLetter(byte value)
        {
            if ((value >= LowercaseA && value <= LowercaseZ) || (value >= UppercaseA && value <= UppercaseZ))
            {
                return true;
            }

            return false;
        }


        /// <summary>
        ///     Returns true if the character is a numeric character
        /// </summary>
        /// <param name="value">The ascii character value to check.</param>
        /// <returns>true if the ascii character is a numeric character, false otherwise.</returns>

        public static bool IsNumeric(byte value)
        {
            return value >= Zero && value <= Nine;
        }


        /// <summary>
        ///     Returns true if the character is a letter or a number
        /// </summary>
        /// <param name="value">The ascii character value to check.</param>
        /// <returns>true if the ascii character is a letter or numeric, false otherwise.</returns>

        public static bool IsLetterOrDigit(byte value)
        {
            if ((value >= LowercaseA && value <= LowercaseZ) || (value >= UppercaseA && value <= UppercaseZ) ||
                (value >= Zero && value <= Nine))
            {
                return true;
            }

            return false;
        }


        /// <summary>
        ///     Returns true if the character is a character used in a hex code
        /// </summary>
        /// <param name="value">The ascii character value to check.</param>
        /// <returns>true if the ascii character can be converted to a hexadecimal value, false otherwise.</returns>

        public static bool IsHexCode(byte value)
        {
            if ((value >= LowercaseA && value <= LowercaseF) || (value >= UppercaseA && value <= UppercaseF) ||
                (value >= Zero && value <= Nine))
            {
                return true;
            }

            return false;
        }


        /// <summary>
        ///     Returns true if the character is a lower cased character
        /// </summary>
        /// <param name="value">The ascii character value to check.</param>
        /// <returns>true if the ascii character is a lower case character, false otherwise.</returns>

        public static bool IsLower(byte value)
        {
            return value >= LowercaseA && value <= LowercaseZ;
        }


        /// <summary>
        ///     Returns true if the character is an upper cased character
        /// </summary>
        /// <param name="value">The ascii character value to check.</param>
        /// <returns>true if the ascii character is an upper case character, false otherwise.</returns>

        public static bool IsUpper(byte value)
        {
            return value >= UppercaseA && value <= UppercaseZ;
        }


        /// <summary>
        ///     Converts the passed upper cased character to its lower case equivalent character.
        /// </summary>
        /// <param name="value">The ascii character value to change to lower case.</param>
        /// <returns>
        ///     A ascii byte value continuing the lower case value of the character, or if there is no lower case character
        ///     then it returns the same value passed in.
        /// </returns>

        public static byte ToLower(byte value)
        {
            if (value >= UppercaseA && value <= UppercaseZ)
            {
                return (byte)(value + UpperLowerLetterOffet);
            }

            return value;
        }


        /// <summary>
        ///     Converts the passed lower cased character to its upper case equivalent character.
        /// </summary>
        /// <param name="value">The ascii character value to change to upper case.</param>
        /// <returns>
        ///     A ascii byte value continuing the upper case value of the character, or if there is no upper case character
        ///     then it returns the same value passed in.
        /// </returns>

        public static byte ToUpper(byte value)
        {
            if (value >= LowercaseA && value <= LowercaseZ)
            {
                return (byte)(value - UpperLowerLetterOffet);
            }

            return value;
        }


        /// <summary>
        ///     Checks if the supplied character represents a control character.
        ///     Between <see cref="ControlCharactersStart">ControlCharactersStart</see> and
        ///     <see cref="ControlCharactersEnd">ControlCharactersEnd</see>, plus the Delete character
        /// </summary>
        /// <param name="value">The ascii character to check.</param>
        /// <returns>true if the ascii character is a control character, false otherwise.</returns>

        public static bool IsControl(byte value)
        {
            return (value <= ControlCharactersEnd) || value == Delete;
        }


        /// <summary>
        ///     Returns if the character represents a symbol character.
        /// </summary>
        /// <param name="value">The ascii character to check.</param>
        /// <returns>true if the ascii character is a symbol character, false otherwise.</returns>
        /// <remarks>
        ///     List of punctuation characters checked
        ///     <list type="bullet">
        ///         <item>
        ///             <term>
        ///                 <see cref="DollarSign">DollarSign</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="Plus">Plus</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="LessThanSign">LessThanSign</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="EqualsSign">EqualsSign</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="GreaterThanSign">GreaterThanSign</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="Caret">Caret</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="SingleQuote">SingleQuote</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="VerticalLine">VerticalLine</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="Tilde">Tilde</see>
        ///             </term>
        ///         </item>
        ///     </list>
        /// </remarks>
        public static bool IsSymbol(byte value)
        {
            switch (value)
            {
                case DollarSign:
                case Plus:
                case LessThanSign:
                case EqualsSign:
                case GreaterThanSign:
                case Caret:
                case SingleQuote:
                case VerticalLine:
                case Tilde:
                    return true;
            }

            return false;
        }


        /// <summary>
        ///     Returns if the char matches a space character.
        /// </summary>
        /// <param name="value">The ascii character to check.</param>
        /// <returns>true if the ascii character is a space character, false otherwise.</returns>

        public static bool IsSeparator(byte value)
        {
            return value == Space;
        }


        /// <summary>
        ///     Checks whether the character is a punctuation character aa.
        /// </summary>
        /// <param name="value">The ascii character to check.</param>
        /// <returns>true if the ascii character is a punctuation character, false otherwise.</returns>
        /// <remarks>
        ///     List of punctuation characters checked
        ///     <list type="bullet">
        ///         <item>
        ///             <term>
        ///                 <see cref="ExclamationMark">ExclamationMark</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="QuotationMark">QuotationMark</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="Crosshatch">CrossHatch</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="PercentSign">PercentSign</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="Ampersand">Ampersand</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="Apostrophe">Apostrophe</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="OpeningParentheses">OpeningParentheses</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="ClosingParentheses">ClosingParentheses</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="Asterisk">Asterisk</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="Comma">Comma</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="Hyphen">Hyphen</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="Period">Period</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="ForwardSlash">ForwardSlash</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="Colon">Colon</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="Semicolon">Semicolon</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="QuestionMark">QuestionMark</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="AtSign">AtSign</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="OpeningSquareBracket">OpeningSquareBracket</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="Backslash">Backslash</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="ClosingSquareBracket">ClosingSquareBracket</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="Underscore">Underscore</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="OpeningCurlyBrace">OpeningCurlyBrace</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="ClosingCurlyBrace">ClosingCurlyBrace</see>
        ///             </term>
        ///         </item>
        ///     </list>
        /// </remarks>
        [SuppressMessage("Microsoft.Usage", "CA2233:OperationsShouldNotOverflow", MessageId = "value-33")]
        [SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        public static bool IsPunctuation(byte value)
        {
            switch (value)
            {
                case ExclamationMark:
                case QuotationMark:
                case Crosshatch:
                case PercentSign:
                case Ampersand:
                case Apostrophe:
                case OpeningParentheses:
                case ClosingParentheses:
                case Asterisk:
                case Comma:
                case Hyphen:
                case Period:
                case ForwardSlash:
                case Colon:
                case Semicolon:
                case QuestionMark:
                case AtSign:
                case OpeningSquareBracket:
                case Backslash:
                case ClosingSquareBracket:
                case Underscore:
                case OpeningCurlyBrace:
                case ClosingCurlyBrace:
                    return true;
            }

            return false;
        }


        /// <summary>
        ///     Check whether the supplied ascii character can be considered as white space.
        /// </summary>
        /// <param name="value">The ascii character to check.</param>
        /// <returns>true if the ascii character is a white space character, false otherwise.</returns>
        /// <remarks>
        ///     List of white space characters checked
        ///     <list type="bullet">
        ///         <item>
        ///             <term>
        ///                 <see cref="HorizontalTab">HorizontalTab</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="Linefeed">LineFeed</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="VerticalTab">VerticalTab</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="FormFeed">FormFeed</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="CharacterReturn">CharacterReturn</see>
        ///             </term>
        ///         </item>
        ///         <item>
        ///             <term>
        ///                 <see cref="Space">Space</see>
        ///             </term>
        ///         </item>
        ///     </list>
        /// </remarks>
        [SuppressMessage("Microsoft.Usage", "CA2233:OperationsShouldNotOverflow", MessageId = "value-9")]
        public static bool IsWhiteSpace(byte value)
        {
            switch (value)
            {
                case HorizontalTab:
                case Linefeed:
                case VerticalTab:
                case FormFeed:
                case CharacterReturn:
                case Space:
                    return true;
            }

            return false;
        }


        /// <summary>
        ///     Converts the passed character byte to a <see cref="char">Char</see> Object
        /// </summary>
        /// <param name="value">The ascii character to convert.</param>
        /// <returns>The ascii character converted to a <see cref="char">Char</see> type.</returns>

        public static char ToChar(byte value)
        {
            return Convert.ToChar(value);
        }


        /// <summary>
        ///     Converts the passed <see cref="char">Char</see> to a character byte
        /// </summary>
        /// <param name="value">The ascii character to convert.</param>
        /// <returns>The ascii character converted to a <see cref="char">Char</see> type.</returns>

        public static byte ToByte(char value)
        {
            return Convert.ToByte(value);
        }


        /// <summary>
        ///     Converts the given ASCII character (as a Byte value) to its Int32 equivalent
        /// </summary>
        /// <param name="value">The ascii character to convert.</param>
        /// <returns>The numeric representation of the ascii character.</returns>

        public static int ToNumber(byte value)
        {
            if (value >= Zero && value <= Nine)
            {
                return value - Zero;
            }

            throw new ArgumentException("Character passed does not represent a valid numeric character");
        }


        /// <summary>
        ///     Converts the given ASCII character (as a Byte value) to an Int32 representing the hexadecimal code.
        /// </summary>
        /// <param name="value">The ascii character to convert.</param>
        /// <returns>A numeric value representing the hexadecimal value of the ascii character passed.</returns>
        /// <remarks>
        ///     Will throw an <see cref="ArgumentOutOfRangeException">ArgumentOutOfRangeException</see> if the value passed
        ///     cannot be converted to a hexadecimal value.
        /// </remarks>
        [SuppressMessage("Microsoft.Usage", "CA2233:OperationsShouldNotOverflow")]
        public static int ToHex(byte value)
        {
            if (value >= Zero && value <= Nine)
            {
                return value - Zero;
            }

            if (value >= LowercaseA && value <= LowercaseF) //a-f = 10-15
            {
                return value - LowercaseA + HexValue;
            }

            if (value >= UppercaseA && value <= UppercaseF) //A-F = 10-15
            {
                return value + HexValue - UppercaseA + HexValue;
            }

            throw new ArgumentOutOfRangeException(nameof(value));
        }
    }
}