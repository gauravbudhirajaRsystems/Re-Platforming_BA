// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace LevitJames.TextServices
{
    /// <summary>
    ///     Builds true rtf from simple formatted source string
    /// </summary>
    public class RichTextBuilder
    {
        private string[] _args;
        private bool _bulleted;
        private List<string> _colors;
        private int _counter;
        private bool _indented;
        private string _pattern;
        private Stack<string> _stack;

        /// <summary>
        ///     Return true if the source string starts with a [rtf] or {rtf} tag.
        /// </summary>
        /// <param name="source"></param>

        public static bool IsSimpleRtf(string source)
        {
            return !string.IsNullOrEmpty(source)
                   && (source.StartsWith("[rtf]", StringComparison.OrdinalIgnoreCase)
                       || source.StartsWith("{rtf}", StringComparison.OrdinalIgnoreCase));
        }


        /// <summary>
        ///     <para>Formats simple string literals to rtf formatted text. See the remarks for a list of supported tags.</para>
        ///     <para>
        ///         Up to 10 numeric tags can be supplied. The numeric tags are in the format {n} where n is the number between 0
        ///         and 9.
        ///     </para>
        /// </summary>
        /// <param name="source">The string for format</param>
        /// <param name="args">
        ///     String literals [0] to [9] will be replaced by the equivalent indexed argument contained in the args
        ///     param array.
        /// </param>
        /// <remarks>
        ///     <list type="bullet">
        ///         <item>[cr{0-9}] Character return. [0-9] is the optional number of returns</item>
        ///         <item>[par]new paragraph.</item>
        ///         <item>[tab{0-9}] Tab character.[0-9] is the optional number of tabs</item>
        ///         <item>[b] Begin Bold run of characters</item>
        ///         <item>[i] Begin Italic run of characters</item>
        ///         <item>[-] Ends a run of characters</item>
        ///         <item>[indent{0-9}] indents a paragraph [0-9] is the optional number of indenting</item>
        ///         <item>[bullet]adds a bullet point to the paragraph</item>
        ///         <item>
        ///             [link (urllink)]Link Text[-] Inserts a hyperlink (urllink) is the link you add. Link Text is the text
        ///             displayed.
        ///             <para>Example: [link http://www.google.com]Search[-]</para>
        ///         </item>
        ///         <item>
        ///             [color {Color In Hex}] A run of colored text[-]. {Color In Hex} is a 6 digit hex value in the format RGB.
        ///             Like so FF00BB
        ///         </item>
        ///     </list>
        /// </remarks>
        /// <returns>A properly formatted rtf string.</returns>
        /// <example>
        ///     FormatRtf("The [color FF0000[[0][-] Fox jumped over the [1]","Red", "Cat") would return the string an rtf formatted
        ///     string where the word "Red" would be formatted as red text when displayed in an RichEdit control.
        /// </example>
        public string Build(string source, string[] args)
        {
            if (string.IsNullOrEmpty(source))
                return string.Empty;

            if (source.StartsWith("[rtf]", StringComparison.OrdinalIgnoreCase))
                return BuildNew(source, args);

            return new RichTextBuilderOld().Build(source, args);
        }

        private string BuildNew(string source, string[] args)
        {
            const string rtfHeaderStart = "{\\rtf1\\ansi{\\colortbl ;\\red0\\green0\\blue255;";
            const string rtfHeaderEnd = "}\\viewkind4\\pard ";

            _colors = null;
            _bulleted = false;
            _indented = false;
            _counter = 0;
            _stack = new Stack<string>();
            _args = args;

            _pattern =
                @"\\|\[cr\]|\[cr\d+\]|\[crlf\d+\]|\[crlf\]|\[par\]|\[tab\]|\[tab\d+\]|\[b\]|\[i\]|\[\-\]|\[indent\]|\[indent\d+\]|\[bullet\]|\[link ([^\\].*?)\]|\r\n|\[color ([^\\].*?)\]";

            //only match for {0} to {9} if we have valid args 
            if (_args != null && _args.Length > 0)
                _pattern += "|\\[[0-9]\\]";

            var rtf = Regex.Replace(source, _pattern, RtfMatch);

            var rtfHeader = new StringBuilder(rtfHeaderStart);

            if (_colors != null && _colors.Count > 1)
            {
                //Start at index 1 to skip blue hyperlink color which is already in the header
                for (var i = 1; i <= _colors.Count - 1; i++)
                {
                    rtfHeader.Append(_colors[i]);
                }
            }

            rtfHeader.Append(rtfHeaderEnd);

            rtfHeader.Append(rtf.StartsWith("[rtf]", StringComparison.OrdinalIgnoreCase)
                                 ? rtf.Substring(startIndex: 5)
                                 : rtf);

            rtfHeader.Append("}");

            return rtfHeader.ToString();
        }


        private string RtfMatch(Match m)
        {
            var value = m.Value.ToUpperInvariant();
            switch (value)
            {
                case @"\":
                    return @"\\";

                case "[B]":
                    _stack.Push(value);
                    return @"\b ";

                case "[I]":
                    _stack.Push(value);
                    return @"\i ";

                case "[U]":
                    _stack.Push(value);
                    return @"\u ";

                case "[PAR]":
                    return @"\par\pard ";
                case "[-]":
                    if (_stack.Count == 0)
                    {
                        return null;
                    }

                    switch (_stack.Pop())
                    {
                        case "[B]":
                            return @"\b0 ";

                        case "[I]":
                            return @"\i0 ";

                        case "[U]":
                            return @"\u0 ";

                        case "[COLOR]":
                            return @"\cf0 ";

                        case "[LINK]":
                            return "}}}";

                        case "[INDENT]":
                        case "[BULLET]":
                            return @"\par\pard ";
                    }

                    break;

                case "[BULLET]":
                    _bulleted = true;
                    return @"\pard{\pntext\fs36\'B7\tab}{\*\pn\pnlvlblt\pnf1\pnindent0{\pntxtb\'B7}}\fi-360\li720 ";

                case "[0]":
                    return Regex.Replace(_args[0], _pattern, RtfMatch);

                case "[1]":
                    return Regex.Replace(_args[1], _pattern, RtfMatch);

                case "[2]":
                    return Regex.Replace(_args[2], _pattern, RtfMatch);

                case "[3]":
                    return Regex.Replace(_args[3], _pattern, RtfMatch);

                case "[4]":
                    return Regex.Replace(_args[4], _pattern, RtfMatch);

                case "[5]":
                    return Regex.Replace(_args[5], _pattern, RtfMatch);

                case "[6]":
                    return Regex.Replace(_args[6], _pattern, RtfMatch);

                case "[7]":
                    return Regex.Replace(_args[7], _pattern, RtfMatch);

                case "[8]":
                    return Regex.Replace(_args[8], _pattern, RtfMatch);

                case "[9]":
                    return Regex.Replace(_args[9], _pattern, RtfMatch);

                default:
                    string rtf;
                    if (ParseCharacterReturns(value, out rtf))
                        return rtf;

                    if (ParseTabs(value, out rtf))
                        return rtf;

                    if (ParseIndents(value, out rtf))
                        return rtf;

                    if (ParseLinks(value, out rtf))
                        return rtf;

                    if (ParseColors(value, out rtf))
                        return rtf;

                    break;
            }

            return null;
        }


        private bool ParseCharacterReturns(string value, out string rtf)
        {
            if (!value.StartsWith("[CR", StringComparison.Ordinal) && (value != Environment.NewLine))
            {
                rtf = null;
                return false;
            }

            if (_indented || _bulleted)
            {
                _bulleted = false;
                _indented = false;
            }

            if (value == "[CR]" || value == "[CR1]" || value == "[CRLF]" || value == "[CRLF1]" ||
                value == Environment.NewLine)
            {
                rtf = @"\par ";
                return true;
            }

            if (value == "[CR2]" || value == "[CRLF2]")
            {
                rtf = @"\par\par ";
                return true;
            }

            _counter = value.StartsWith("[CRLF", StringComparison.Ordinal) ? 5 : 3;
            var s = value.Substring(_counter, value.Length - (_counter + 1));
            if (int.TryParse(s, out _counter))
                _counter = 1;

            if (_counter < 1)
                _counter = 1;

            rtf = string.Join(separator: null, values: Enumerable.Repeat(@"\par", _counter)) + " ";
            return true;
        }


        private bool ParseTabs(string value, out string rtf)
        {
            if (!value.StartsWith("[TAB", StringComparison.Ordinal) && (value != Convert.ToChar(value: 9).ToString()))
            {
                rtf = null;
                return false;
            }

            _indented = true;

            if (value == "[TAB]" || value == "[TAB1]" || value == Convert.ToChar(value: 9).ToString())
            {
                rtf = @"\tab ";
                return true;
            }

            if (value == "[TAB2]")
            {
                rtf = @"\tab\tab ";
                return true;
            }


            if (!int.TryParse(value.Substring(startIndex: 4, length: value.Length - 5), out _counter))
                _counter = 1;

            if (_counter < 1)
                _counter = 1;

            rtf = string.Join(separator: null, values: Enumerable.Repeat(@"\tab", _counter)) + " ";
            return true;
        }


        private bool ParseIndents(string value, out string rtf)
        {
            if (!value.StartsWith("[INDENT", StringComparison.Ordinal))
            {
                rtf = null;
                return false;
            }

            if (!int.TryParse(value.Substring(startIndex: 7, length: value.Length - 8), out _counter))
                _counter = 1;

            if (_counter < 1)
                _counter = 1;

            _indented = true;
            rtf = @"\li" + (360 * _counter).ToString(CultureInfo.InvariantCulture) + " ";
            return true;
        }


        private bool ParseLinks(string value, out string rtf)
        {
            if (!value.StartsWith("[LINK ", StringComparison.Ordinal))
            {
                rtf = null;
                return false;
            }

            //Make link to lower and replace any \ chars with / chars.
            var link = value.Substring(startIndex: 6, length: value.Length - 7)
                            .ToLower()
                            .Replace(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            _stack.Push("[LINK]");
            rtf = "{\\field{\\*\\fldinst{HYPERLINK \"" + link + "\"}}{\\fldrslt{\\cf1\\ul ";
            return true;
        }


        private bool ParseColors(string value, out string rtf)
        {
            if (!value.StartsWith("[COLOR ", StringComparison.Ordinal))
            {
                rtf = null;
                return false;
            }

            rtf = null;
            _stack.Push("[COLOR]");
            var color = value.Substring(startIndex: 6, length: value.Length - 7).TrimStart();
            if (TryParseHexStringToRtfColor(color, out color))
            {
                var colors = _colors;
                _colors = colors ?? new List<string> { @"\red0\green0\blue255;" };
                var clrIndex = _colors.IndexOf(color) + 1;
                if (clrIndex == 0)
                {
                    _colors.Add(color);
                    clrIndex = _colors.Count;
                }

                rtf = @"\cf" + clrIndex.ToString(CultureInfo.InvariantCulture);
            }

            return true;
        }


        /// <summary>
        ///     Parses a hex color value into a
        /// </summary>
        /// <param name="hexColor"></param>
        /// <param name="rtfColor"></param>
        private static bool TryParseHexStringToRtfColor(string hexColor, out string rtfColor)
        {
            rtfColor = null;

            if (hexColor == null || hexColor.Length != 6)
            {
                // you can choose whether to throw an exception
                //throw new ArgumentException("hexColor is not exactly 6 digits.");
                return false;
            }

            try
            {
                rtfColor = "\\red" +
                           int.Parse(hexColor.Substring(startIndex: 0, length: 2), NumberStyles.HexNumber,
                                     CultureInfo.CurrentCulture) + "\\green" +
                           int.Parse(hexColor.Substring(startIndex: 2, length: 2), NumberStyles.HexNumber,
                                     CultureInfo.CurrentCulture) + "\\blue" +
                           int.Parse(hexColor.Substring(startIndex: 4, length: 2), NumberStyles.HexNumber,
                                     CultureInfo.CurrentCulture) + ";";

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}