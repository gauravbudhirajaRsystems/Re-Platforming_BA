// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace LevitJames.TextServices
{
    /// <summary>
    ///     Builds true rtf from simple formatted source string, Uses curly braces. Now obsolete use RtfBuilder instead
    /// </summary>
    internal class RichTextBuilderOld
    {
        private string[] _args;
        private bool _bulleted;
        private List<string> _colors;
        private bool _indented;
        private string _pattern;
        private Stack<string> _stack;

        public string Build(string source, string[] args)
        {
            const string rtfHeaderStart = "{\\rtf1\\ansi{\\colortbl ;\\red0\\green0\\blue255;";
            const string rtfHeaderEnd = "}\\viewkind4\\pard ";

            _bulleted = false;
            _indented = false;
            _colors = null;
            _stack = new Stack<string>();
            _args = args;

            _pattern =
                @"\\|\{cr\}|\{cr\d+\}|\{crlf\d+\}|\{crlf\}|\{tab\}|\{tab\d+\}|\{b\}|\{i\}|\{\-\}|\{indent\}|\{indent\d+\}|\{bullet\}|\{link ([^\\].*?)\}|\r\n|{color ([^\\].*?)\}";

            if (args != null && args.Length > 0)
                _pattern += "|\\{[0-9]\\}";

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

            rtfHeader.Append(rtf.StartsWith("{rtf}", StringComparison.OrdinalIgnoreCase)
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

                case "{B}":
                    _stack.Push(value);
                    return @"\b ";

                case "{I}":
                    _stack.Push(value);
                    return @"\i ";

                case "{U}":
                    _stack.Push(value);
                    return @"\u ";

                case "{-}":
                    if (_stack.Count == 0)
                    {
                        return null;
                    }

                    switch (_stack.Pop())
                    {
                        case "{B}":
                            return @"\b0 ";

                        case "{I}":
                            return @"\i0 ";

                        case "{U}":
                            return @"\u0 ";

                        case "{COLOR}":
                            return @"\cf0 ";

                        case "{LINK}":
                            return "}}}";

                        case "{INDENT}":
                        case "{BULLET}":
                            return @"\par\pard ";
                    }

                    break;

                case "{BULLET}":
                    _bulleted = true;
                    return @"\pard{\pntext\fs36\'B7\tab}{\*\pn\pnlvlblt\pnf1\pnindent0{\pntxtb\'B7}}\fi-360\li720 ";

                case "{0}":
                    return Regex.Replace(_args[0], _pattern, RtfMatch);

                case "{1}":
                    return Regex.Replace(_args[1], _pattern, RtfMatch);

                case "{2}":
                    return Regex.Replace(_args[2], _pattern, RtfMatch);

                case "{3}":
                    return Regex.Replace(_args[3], _pattern, RtfMatch);

                case "{4}":
                    return Regex.Replace(_args[4], _pattern, RtfMatch);

                case "{5}":
                    return Regex.Replace(_args[5], _pattern, RtfMatch);

                case "{6}":
                    return Regex.Replace(_args[6], _pattern, RtfMatch);

                case "{7}":
                    return Regex.Replace(_args[7], _pattern, RtfMatch);

                case "{8}":
                    return Regex.Replace(_args[8], _pattern, RtfMatch);

                case "{9}":
                    return Regex.Replace(_args[9], _pattern, RtfMatch);

                default:
                    int counter;

                    if (value.StartsWith("{CR", StringComparison.Ordinal) || value == Environment.NewLine)
                    {
                        if (_indented || _bulleted)
                        {
                            _bulleted = false;
                            _indented = false;
                        }

                        if (value == "{CRr}" || value == "{CR1}" || value == "{CRLF}" || value == "{CRLF1}" ||
                            value == Environment.NewLine)
                            return @"\par\pard ";

                        if (value == "{CR2}" || value == "{CRLF2}")
                            return @"\par\pard\par ";

                        counter = value.StartsWith("{CRLF", StringComparison.Ordinal) ? 5 : 3;
                        var s = value.Substring(counter, value.Length - (counter + 1));
                        if (int.TryParse(s, out counter))
                            counter = 1;

                        if (counter < 1)
                            counter = 1;

                        return string.Join(separator: null, values: Enumerable.Repeat(@"\par", counter)) + " ";
                    }

                    if (value.StartsWith("{TAB", StringComparison.Ordinal) ||
                        value == Convert.ToChar(value: 9).ToString())
                    {
                        _indented = true;

                        if (value == "{TAB}" || value == "{TAB1}" ||
                            value == Convert.ToChar(value: 9).ToString())
                            return @"\t ";

                        if (value == "{TAB2}")
                            return @"\t\t ";

                        if (!int.TryParse(value.Substring(startIndex: 4, length: value.Length - 5), out counter))
                            counter = 1;

                        if (counter < 1)
                            counter = 1;

                        return string.Join(separator: null, values: Enumerable.Repeat(@"\t", counter)) + " ";
                    }

                    if (value.StartsWith("{INDENT", StringComparison.Ordinal))
                    {
                        if (!int.TryParse(value.Substring(startIndex: 7, length: value.Length - 8), out counter))
                            counter = 1;

                        if (counter < 1)
                            counter = 1;

                        _indented = true;
                        return @"\li" + (360 * counter).ToString(CultureInfo.InvariantCulture) +
                               " ";
                    }

                    if (value.StartsWith("{LINK ", StringComparison.Ordinal))
                    {
                        _stack.Push("{LINK}");
                        return "{\\field{\\*\\fldinst{HYPERLINK \"" +
                               value.Substring(startIndex: 6, length: value.Length - 7) +
                               "\"}}{\\fldrslt{\\cf1\\ul ";
                    }

                    if (value.StartsWith("{COLOR ", StringComparison.Ordinal))
                    {
                        _stack.Push("{COLOR}");
                        var color = value.Substring(startIndex: 6, length: value.Length - 7).TrimStart();
                        if (TryParseHexStringToRtfColor(color, out color))
                        {
                            if (_colors == null)
                                _colors = new List<string> { @"\red0\green0\blue255;" };
                            var clrIndex = _colors.IndexOf(color) + 1;
                            if (clrIndex == 0)
                            {
                                _colors.Add(color);
                                clrIndex = _colors.Count;
                            }

                            return @"\cf" + clrIndex.ToString(CultureInfo.InvariantCulture);
                        }
                    }

                    break;
            }

            return null;
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