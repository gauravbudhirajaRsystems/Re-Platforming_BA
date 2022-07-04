// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections.Generic;
using System.Data.Entity.Design.PluralizationServices;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using JetBrains.Annotations;
using LevitJames.Core;

namespace LevitJames.TextServices
{
    public static partial class CompilerExtensions
    {
        private static readonly char[] HyphenChars = {
                                                         (char) 0x002D, //45,
                                                         (char) 0x001E, //30,
                                                         (char) 0x001F, //31,
                                                         (char) 0x0096, //150,
                                                         (char) 0x0097, //151,
                                                         (char) 0x2010, //8208,
                                                         (char) 0x2011, //8209,
                                                         (char) 0x2012, //8210,
                                                         (char) 0x2013, //8211,
                                                         (char) 0x2014, //8212,
                                                         (char) 0x2015 //8213
                                                     };


        /// <summary>
        ///     Returns an array of plain quote characters. Currently characters "(34) or '(39)
        /// </summary>
        private static readonly char[] PlainQuoteChars = {
                                                             (char) 0x0022, // 34,
                                                             (char) 0x0027 // 39
                                                         };


        /// <summary>
        ///     Returns an array of characters representing smart quotes.
        /// </summary>
        /// <remarks>
        ///     <list type="char">
        ///         <value>145</value>
        ///         <value>146</value>
        ///         <value>147</value>
        ///         <value>148</value>
        ///         <value>8217</value>
        ///         <value>8218</value>
        ///         <value>8219</value>
        ///         <value>8220</value>
        ///         <value>8221</value>
        ///     </list>
        /// </remarks>
        private static readonly char[] SmartQuoteChars = {
                                                             (char) 0x0091, // 145,
                                                             (char) 0x0092, // 146,
                                                             (char) 0x0093, // 147,
                                                             (char) 0x0094, // 148,
                                                             (char) 0x2018, // 8216
                                                             (char) 0x2019, // 8217
                                                             (char) 0x201B, // 8219 
                                                             (char) 0x201C, // 8220 
                                                             (char) 0x201D // 8221 
                                                         };


        /// <summary>
        ///     Returns an array of characters representing white space in a string.
        /// </summary>
        private static readonly char[] WhiteSpaceChars = {
                                                             (char) 0x0009, //9, character tabulation
                                                             (char) 0x000A, //10, line feed
                                                             (char) 0x000B, //11, line tabulation
                                                             (char) 0x000C, //12, form feed
                                                             (char) 0x000D, //13, carriage return
                                                             (char) 0x0020, //32, space
                                                             (char) 0x0085, //133, next line
                                                             (char) 0x00A0, //160, no-break space
                                                             (char) 0x1680, //5760, ogham space mark
                                                             (char) 0x2000, //8192, en quad
                                                             (char) 0x2001, //8193, em quad
                                                             (char) 0x2002, //8194, en space
                                                             (char) 0x2003, //8195, em space
                                                             (char) 0x2004, //8196, three-per-em space
                                                             (char) 0x2005, //8197, four-per-em space
                                                             (char) 0x2006, //8198, six-per-em space
                                                             (char) 0x2007, //8199, figure space
                                                             (char) 0x2008, //8200, punctuation space
                                                             (char) 0x2009, //8201, thin space
                                                             (char) 0x200A, //8202, hair space
                                                             (char) 0x2028, //8232, line separator
                                                             (char) 0x2029, //8233, paragraph separator
                                                             (char) 0x202F, //8239, narrow no-break space
                                                             (char) 0x205F, //8287, medium mathematical space
                                                             (char) 0x3000 //12288, ideographic space
                                                         };


        /// <summary>
        ///     Returns an array of characters representing spaces in a string.
        /// </summary>
        private static readonly char[] SpaceChars = {
                                                        (char) 0x0020, //32, space
                                                        (char) 0x00A0, //160, no-break space
                                                        (char) 0x1680, //5760, ogham space mark
                                                        (char) 0x2000, //8192, en quad
                                                        (char) 0x2001, //8193, em quad
                                                        (char) 0x2002, //8194, en space
                                                        (char) 0x2003, //8195, em space
                                                        (char) 0x2004, //8196, three-per-em space
                                                        (char) 0x2005, //8197, four-per-em space
                                                        (char) 0x2006, //8198, six-per-em space
                                                        (char) 0x2007, //8199, figure space
                                                        (char) 0x2008, //8200, punctuation space
                                                        (char) 0x2009, //8201, thin space
                                                        (char) 0x200A, //8202, hair space
                                                        (char) 0x202F, //8239, narrow no-break space
                                                        (char) 0x205F, //8287, medium mathematical space
                                                        (char) 0x3000 //12288, ideographic space
                                                    };

        private static readonly char[] NoWidthSpecialChars =
        {
            (char) 0x200C, //8204 Zero-width non-joiner (No-Width Optional Break)
            (char) 0x200D, //8205 Zero-width joiner (No-Width Non Break)
            (char) 0x200E, //8206 Left-to-Right Mark
            (char) 0x200F, //8207 Right-to-Left Mark
            (char) 0x202A, //8234 Left-to-Right Embedding
            (char) 0x202B, //8235 Right-to-Left Embedding
            (char) 0x202C, //8236 Pop Directional Formatting
            (char) 0x202D, //8237 Left-to-Right Override
            (char) 0x202E //8238 Right-to-Left Override
        };


        /// <summary>Splits a camel cased string my adding a space where an upper case letter is found.</summary>
        /// <param name="source"></param>
        public static string ToSpacedCase(this string source)
        {
            return ToSpacedCase(source, trimStart: null, trimEnd: null, style: SpacedCaseStyle.CapitalizeFirstWordOnly,
                                ignoreCase: true);
        }

        /// <summary>Splits a camel cased string by adding a space where an upper case letter is found.</summary>
        /// <param name="source"></param>
        /// <param name="style"></param>
        public static string ToSpacedCase(this string source, SpacedCaseStyle style)
        {
            return ToSpacedCase(source, trimStart: null, trimEnd: null, style: style, ignoreCase: true);
        }

        /// <summary>Splits a camel cased string my adding a space where an upper case letter is found.</summary>
        /// <param name="source">The source string to convert.</param>
        /// <param name="trimStart">the text to remove from the start of the source string.</param>
        /// <param name="trimEnd">the text to remove from the end of the source string.</param>
        /// <param name="ignoreCase">Whether to ignore the case when trimming the start and end strings.</param>
        /// <param name="style">
        ///     One of the SpacedCaseStyle Enum members, used to define the capitalization style of the output
        ///     string.
        /// </param>
        [SuppressMessage("Microsoft.Globalization", "CA1307:SpecifyStringComparison",
            MessageId = "System.String.StartsWith(System.String)")]
        public static string ToSpacedCase(this string source, string trimStart, string trimEnd, bool ignoreCase,
                                          SpacedCaseStyle style)
        {
            if (string.IsNullOrWhiteSpace(source))
                return source;
            if (!string.IsNullOrEmpty(trimStart) &&
                source.StartsWith(trimStart, ignoreCase, CultureInfo.CurrentUICulture))
            {
                source = source.Substring(trimStart.Length);
            }

            if (!string.IsNullOrEmpty(trimEnd) &&
                source.EndsWith(trimEnd, ignoreCase, CultureInfo.CurrentUICulture))
            {
                source = source.Substring(startIndex: 0, length: source.Length - trimEnd.Length);
            }

            //var str = Regex.Replace(source, "([A-Z])", " $1", RegexOptions.Compiled).TrimStart();
            //Note this regex statement keeps achonims such as BASomeText from becoming Ba Some Text
            var str = Regex.Replace(Regex.Replace(source, @"(\P{Ll})(\P{Ll}\p{Ll})", "$1 $2"), @"(\p{Ll})(\P{Ll})", "$1 $2").TrimStart();

            // We've trimmed the start of str to get rid of false extra space introduced in the regex replace.
            // However, source could also start with an arbitrary number of spaces. 
            // Make sure str starts with the same number of spaces that source does

            var spaceCount = 0;
            foreach (var c in source)
            {
                if (c != ' ')
                    break;
                spaceCount += 1;
            }

            str = new string(' ', spaceCount) + str;

            var retVal = string.Empty;

            switch (style)
            {
                case SpacedCaseStyle.None:
                    retVal = str;
                    break;
                case SpacedCaseStyle.CapitalizeFirstWordOnly:
                    retVal += str.Substring(spaceCount, 1).ToUpper() + str.Substring(spaceCount + 1).ToLower();
                    break;
                case SpacedCaseStyle.CapitalizeAllWords:
                    retVal = Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(str);
                    break;
            }

            return retVal;
        }


        /// <summary>
        ///     Converts the string to Title case using the Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase method.
        ///     Words that are entirely uppercase are not changed and are considered acronyms.
        /// </summary>
        /// <param name="source">The source string to convert to Title case</param>
        public static string ToTitleCase(this string source)
        {
            return source == null ? null : Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(source);
        }


        /// <summary>
        ///     <para>Ensures that a string value is unique by appending an index counter to the string.</para>
        ///     <para>This method can be used to create unique strings used when creating a copy of an item.</para>
        /// </summary>
        /// <param name="source">The string to ensure is unique.</param>
        /// <param name="exists">A predicate that returns if the string is unique or not.</param>
        /// <param name="format">
        ///     The format of the string to append if it is not unique. {0} is the source string, {1} is where
        ///     the index counter is placed and {2} is the optional string suffix,
        /// </param>
        /// <param name="suffix">A string to add to the end of the source string, such as a file extension</param>
        /// <param name="checkBareSourceName">true to first check if the an item with the supplied source and prefix exists without any numeric counter.</param>
        /// <returns>The new string.</returns>
        /// <example>
        ///     <code title="Example" description="" lang="C#">
        /// var dict = new Dictionary&lt;string, string&gt;() {{"Note", "SomeData"}};
        ///  
        /// var newItem = "Note".EnsureUniqueString((key) =&gt; dict.ContainsKey(key));
        /// dict.Add(newItem, dict["Note"]); // newItem will be set to "Note - Copy (1)"
        ///  
        /// newItem = "Note".EnsureUniqueString((key) =&gt; dict.ContainsKey(key));
        /// dict.Add(newItem, dict["Note"]); // newItem will be set to "Note - Copy (2)"</code>
        /// </example>
        public static string EnsureUniqueString(this string source, Predicate<string> exists,
                                                string format = "{0}{2} - Copy ({1})", string suffix = null, bool checkBareSourceName = true)
        {
            if (exists == null)
                throw new ArgumentNullException(nameof(exists));
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            var newString = string.Format(format, source, null, suffix);
            if (checkBareSourceName && exists(newString) == false)
                return source;

            uint nKey = 1;

            do
            {

                newString = string.Format(CultureInfo.CurrentCulture, format, source, nKey, suffix);
                if (nKey > int.MaxValue)
                {
                    //If the start key was not changed i.e. it did not contain any numbers then
                    source += " " + source;
                    nKey = 0;
                }

                if (exists(newString) == false)
                    return newString;

                nKey += 1u;
            } while (true);
        }


        /// <summary>
        ///     Formats simple string literals, such as [crlf] or [tab] to their string equivalents.
        ///     <para>
        ///         If the string starts with [rtf] then the string will be formatted using the FormatRtf method. See the remarks
        ///         section for the list of available literals.
        ///         See the FormatRtf
        ///     </para>
        /// </summary>
        /// <param name="source">The string for format</param>
        /// <remarks>
        ///     <list type="bullet">
        ///         <item>[cr[0-9]] Character return. [0-9] is the optional number of returns</item>
        ///         <item>[crlf[0-9]] Character return. [0-9] is the optional number of returns</item>
        ///         <item>[lf[0-9]] Character return. [0-9] is the optional number of returns</item>
        ///         <item>{[tab[0-9]] Tab character.[0-9] is the optional number of tabs</item>
        ///     </list>
        /// </remarks>
        /// <returns>A properly formatted string.</returns>
        [Obsolete("Do not use")]
        public static string FormatLiterals(this string source)
        {
            return FormatLiterals(source, args: null);
        }

        /// <summary>Formats simple string literals, such as [crlf] or [tab] to there string equivalents.</summary>
        /// <param name="source">The string for format</param>
        /// <param name="args">
        ///     String literals [0] to [9] will be replaced by the equivalent indexed argument contained in the
        ///     args param array.
        /// </param>
        /// <remarks>
        ///     <para>If the string starts with [rtf] then the string will be formatted using the FormatRtf method.</para>
        ///     <list type="bullet">
        ///         <item>[cr[0-9]] Character return. [0-9] is the optional number of returns</item>
        ///         <item>[crlf[0-9]] Character return. [0-9] is the optional number of returns</item>
        ///         <item>[lf[0-9]] Character return. [0-9] is the optional number of line feeds</item>
        ///         <item>{tab[0-9]] Tab character.[0-9] is the optional number of tabs</item>
        ///     </list>
        /// </remarks>
        /// <returns>A formatted string.</returns>
        [Obsolete("Do not use")]
        public static string FormatLiterals(this string source, params string[] args)
        {
            if (string.IsNullOrEmpty(source))
                return source;

            if (RichTextBuilder.IsSimpleRtf(source))
            {
                return FormatRtf(source, args);
            }

            var pattern = "\\[cr\\]|\\[crlf\\]|\\[tab\\]";
            //var pattern = ApplicationHelper.UseNewRtfBuilder ? "\\[cr\\]|\\[crlf\\]|\\[tab\\]" : "\\{cr\\}|\\{crlf\\}|\\{tab\\}";

            if (args != null && args.Length > 0)
                pattern += "|\\[[0-9]\\]";

            //pattern += ApplicationHelper.UseNewRtfBuilder ? "|\\[[0-9]\\]": "|\\{[0-9]\\}";

            //only match for {0} to {9} if we have valid args 
            return Regex.Replace(source, pattern, FormatLiteralsMatchNew(args));

            //return Regex.Replace(source, pattern, ApplicationHelper.UseNewRtfBuilder 
            //    ? FormatLiteralsMatchNew(args)
            //    : FormatLiteralsMatchOld(args));
        }

        private static MatchEvaluator FormatLiteralsMatchNew(string[] args)
        {
            return m =>
                   {
                       var value = m.Value.ToUpperInvariant();
                       switch (value)
                       {
                           case "[CR]":
                           case "[CR1]":
                           case "[CRLF]":
                           case "[CRLF1]":
                               return Environment.NewLine;
                           case "[TAB]":
                           case "[TAB1]":
                               return Convert.ToChar(value: 9).ToString();
                           case "[LF]":
                               return Environment.NewLine;
                           case "[0]":
                               return args?[0];
                           case "[1]":
                               return args?[1];
                           case "[2]":
                               return args?[2];
                           case "[3]":
                               return args?[3];
                           case "[4]":
                               return args?[4];
                           case "[5]":
                               return args?[5];
                           case "[6]":
                               return args?[6];
                           case "[7]":
                               return args?[7];
                           case "[8]":
                               return args?[8];
                           case "[9]":
                               return args?[9];
                           default:
                               {
                                   if (TryParseRtfCounterToken(value, "[CRLF", Convert.ToChar(value: 9).ToString(), out value))
                                       return value;
                                   if (TryParseRtfCounterToken(value, "[CR", Convert.ToChar(value: 9).ToString(), out value))
                                       return value;
                                   if (TryParseRtfCounterToken(value, "[TAB", Convert.ToChar(value: 9).ToString(), out value))
                                       return value;

                                   break;
                               }
                       }

                       return null;
                   };
        }

#if false // for reference
        private static MatchEvaluator FormatLiteralsMatchOld(string[] args)
		{
			return m =>
			{
				var value = m.Value.ToUpperInvariant();
				switch (value)
				{
					case "{CR}":
					case "{CR1}":
					case "{CRLF}":
					case "{CRLF1}":
						return Environment.NewLine;
					case "{TAB}":
					case "{TAB1}":
						return Convert.ToChar(value: 9).ToString();
					case "{LF}":
						return Environment.NewLine;
					case "{0}":
						return args?[0];
					case "{1}":
						return args?[1];
					case "{2}":
						return args?[2];
					case "{3}":
						return args?[3];
					case "{4}":
						return args?[4];
					case "{5}":
						return args?[5];
					case "{6}":
						return args?[6];
					case "{7}":
						return args?[7];
					case "{8}":
						return args?[8];
					case "{9}":
						return args?[9];
					default:
						{
							if (TryParseRtfCounterToken(value, "{CRLF",Convert.ToChar(value: 9).ToString(), out value))
								return value;
							if (TryParseRtfCounterToken(value, "{CR",Convert.ToChar(value: 9).ToString(), out value))
								return value;
							if (TryParseRtfCounterToken(value, "{TAB",Convert.ToChar(value: 9).ToString(), out value))
								return value;

							break;
						}
				}
				return null;
			};
		}
#endif


        private static bool TryParseRtfCounterToken(string value, string token, string parsedToken, out string result)
        {
            result = null;

            if (!(value.StartsWith(token, StringComparison.CurrentCulture) ||
                  value == Convert.ToChar(value: 9).ToString()))
                return false;

            if (int.TryParse(value.Substring(token.Length, value.Length - token.Length), out int counter) == false ||
                counter == 0)
                counter = 1;

            result = string.Join(separator: null, values: Enumerable.Repeat(parsedToken, counter));
            return true;
        }


        /// <summary>
        ///     Formats simple string literals to proper rtf
        /// </summary>
        /// <param name="source">The string for format</param>
        /// <remarks>
        ///     <list type="bullet">
        ///         <item>[cr[0-9]} Character return. [0-9] is the optional number of returns.</item>
        ///         <item>[tab[0-9]] Tab character.[0-9] is the optional number of tabs.</item>
        ///         <item>[b] Begin Bold run of characters.</item>
        ///         <item>[i] Begin Italic run of characters.</item>
        ///         <item>[-] Ends a run of characters.</item>
        ///         <item>[indent[0-9]] indents a paragraph [0-9] is the optional number of indenting.</item>
        ///         <item>[bullet] adds a bullet point to the paragraph.</item>
        ///         <item>
        ///             [link [urllink]]Link Text[-] Inserts a hyperlink [urllink] is the link you add. Link Text is the text
        ///             displayed.
        ///             <para>Example: [link http://www.google.com]Search[-]</para>
        ///         </item>
        ///         <item>
        ///             [color [Color In Hex]] A run of colored text[-]. [Color In Hex] is a 6 digit hex value in the format RGB.
        ///             Like so FF00BB
        ///         </item>
        ///     </list>
        /// </remarks>
        /// <returns>A properly formatted rtf string.</returns>
        public static string FormatRtf(this string source)
        {
            if (!RichTextBuilder.IsSimpleRtf(source))
                return source;
            return FormatRtf(source, args: null);
        }


        /// <summary>
        ///     <para>Formats simple string literals to rtf formatted text. See the remarks for a list of supported tags.</para>
        ///     <para>
        ///         Up to 10 numeric tags can be supplied. The numeric tags are in the format [n] where n is the number between 0
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
        ///         <item>[cr[0-9]] Character return. [0-9] is the optional number of returns</item>
        ///         <item>[tab[0-9]] Tab character.[0-9] is the optional number of tabs</item>
        ///         <item>[b] Begin Bold run of characters</item>
        ///         <item>[i] Begin Italic run of characters</item>
        ///         <item>[-] Ends a run of characters</item>
        ///         <item>[indent[0-9]] indents a paragraph [0-9] is the optional number of indenting</item>
        ///         <item>[bullet] adds a bullet point to the paragraph</item>
        ///         <item>
        ///             [link [urllink]]Link Text[-] Inserts a hyperlink [urllink] is the link you add. Link Text is the text
        ///             displayed.
        ///             <para>Example: {link http://www.google.com}Search[-]</para>
        ///         </item>
        ///         <item>
        ///             {color [Color In Hex]] A run of colored text[-]. [Color In Hex] is a 6 digit hex value in the format RGB.
        ///             Like so FF00BB
        ///         </item>
        ///     </list>
        /// </remarks>
        /// <returns>A properly formatted rtf string.</returns>
        /// <example>
        ///     FormatRtf("The [color FF0000}[0][-] Fox jumped over the [1]","Red", "Cat") would return the string an rtf formatted
        ///     string where the word "Red" would be formatted as red text when displayed in an RichEdit control.
        /// </example>
        public static string FormatRtf(this string source, params string[] args)
        {
            return new RichTextBuilder().Build(source, args);
        }


        /// <summary>Normalizes text, including spaces, hyphens, quotes, and preserving length.</summary>
        /// <param name="source">Text to clean.</param>

        public static string Clean(this string source)
        {
            return source.Clean(true, true, true, false, true);
        }

        /// <summary>Normalizes text.</summary>
        /// <param name="source">Text to clean.</param>
        /// <param name="preserveLength">Set to true to ensure that returned string has same length as input string.</param>
        /// <returns>Normalized text.</returns>
        public static string Clean(this string source, bool preserveLength)
        {
            return source.Clean(true, true, true, preserveLength: preserveLength,
                                trimSpaces: true);
        }

        /// <summary>Normalizes text</summary>
        /// <param name="source">Text to clean.</param>
        /// <param name="doSpaces">Set to true to set all space characters to standard space (ANSI 32).</param>
        /// <param name="doHyphens">Set to true to convert all hyphen characters to standard hyphen (ANSI 30).</param>
        /// <param name="doQuotes">Set to true to convert all smart quotes to straight quotes.</param>
        /// <param name="preserveLength">Set to true to ensure that returned string has same length as input string.</param>
        /// <param name="trimSpaces">Set to true to trim spaces from the start/end of the return text.</param>
        /// <returns>Normalized text.</returns>
        /// <remarks>
        ///     All characters in the source string with ASCII values of 1 through 29 are converted to spaces.
        ///     All double spaces are converted to single spaces.
        ///     Conversion of quotes and hyphens includes the wide character set.
        /// </remarks>
        // See TSWA 7010. Documents with text copied from the internet can have special
        // no-width characters. They can also be inserted via Word>Insert>Symbol>Special Characters.
        //public static string Clean(this string source, bool doSpaces, bool doHyphens, bool doQuotes, bool preserveLength,
        //                           bool trimSpaces, bool doNoWidthSpecialChars)
        public static string Clean(this string source, bool doSpaces, bool doHyphens, bool doQuotes, bool preserveLength,
            bool trimSpaces)
        {
            if (string.IsNullOrEmpty(source))
                return source;

            const char doubleQuote = (char)34;
            const char singleQuote = (char)39;
            const char hyphen = (char)45;
            const char space = (char)32;
            const int loAscii = 30;

            var lastPos = -1;
            var pos = 0;
            var changed = false;
            //var sb = new StringBuilder(source);
            var sb = source.ToCharArray();
            for (var i = 0; i < source.Length; i++)
            {
                var chr = sb[i];
                // All characters ASCII 1 - 29 are converted to spaces
                if (chr < loAscii || doSpaces && chr != space && IsSpace(chr))
                {
                    if (!preserveLength && lastPos == pos)
                        continue; // Ignore if last char was also a space
                    sb[pos++] = space;
                    lastPos = pos;
                    changed = true;
                    continue;
                }

                if (doHyphens && chr != hyphen && IsHyphen(chr))
                {
                    sb[pos++] = hyphen;
                    changed = true;
                    continue;
                }

                if (doQuotes && IsSmartQuote(chr))
                {
                    switch ((int)chr)
                    {
                        case 145:
                        case 146:
                        case 8216:
                        case 8217:
                            sb[pos++] = singleQuote;
                            changed = true;
                            break;
                        default:
                            sb[pos++] = doubleQuote;
                            changed = true;
                            break;
                    }

                    continue;
                }

                // See TSWA 7010. Documents with text copied from the internet can have special
                // no-width characters. They can also be inserted via Word>Insert>Symbol>Special Characters.
                //if (doNoWidthSpecialChars && IsNoWidthSpecialCharacter(chr))
                //{
                //    sb[pos++] = space;
                //    changed = true;
                //    break;
                //}

                if (pos != i)
                {
                    sb[pos] = chr;
                    changed = true;
                }

                pos++;
            }

            var result = changed ? new string(sb, 0, pos) : source;

            // See TSWA 7010. Documents with text copied from the internet can have special
            // no-width characters. They can also be inserted via Word>Insert>Symbol>Special Characters.
            //if (trimSpaces)
            //    result = result.Trim();

            //if (!preserveLength)
            //{
            //    // Remove extraneous spaces
            //    var len = result.Length;
            //    result = DeDupSpaces(result);
            //    while (result.Length != len)
            //    {
            //        len = result.Length;
            //        result = DeDupSpaces(result);
            //    }

            //    string DeDupSpaces(string s) => s.Replace("  ", " ");
            //}

            //return result;

            return trimSpaces ? result.Trim() : result;

            //Old code
            //// Replace low-ASCII values with a space
            //for (var position = 0; position < source.Length; position++)
            //{
            //    var tempChar = source.Substring(position, 1).ToCharArray()[0];
            //    if (tempChar.CompareTo(Convert.ToChar(30)) < 0)
            //    {
            //        // All characters ASCII 1 - 29 are converted to spaces
            //        tempChar = space;
            //    }
            //    result += tempChar;
            //}

            //if (doQuotes)
            //{
            //    result = result.Replace((char)145, singleQuote); // single quote open
            //    result = result.Replace((char)146, singleQuote); // single quote close
            //    result = result.Replace((char)147, doubleQuote); // double quote open
            //    result = result.Replace((char)148, doubleQuote); // double quote close
            //    result = result.Replace((char)8216, singleQuote); // single quote open (wide)
            //    result = result.Replace((char)8217, singleQuote); // single quote close (wide)
            //    result = result.Replace((char)8220, doubleQuote); // double quote open (wide)
            //    result = result.Replace((char)8221, doubleQuote); // double quote close (wide)
            //}


            //if (doHyphens)
            //{
            //    result = result.Replace((char)30, hyphen); // non-breaking hyphen
            //    result = result.Replace((char)31, hyphen); // optional hyphen
            //    result = result.Replace((char)150, hyphen); // en-dash
            //    result = result.Replace((char)151, hyphen); // em-dash
            //    result = result.Replace((char)8211, hyphen); // en-dash (wide)
            //    result = result.Replace((char)8212 hyphen); // em-dash (wide)
            //}

            //if (doSpaces)
            //{
            //    // NOTE: The wide space characters are an ASCII 32 non-wide
            //    result = result.Replace((char)8194), space); // em-space (wide)
            //    result = result.Replace((char)8195), space); // en-space (wide)
            //    result = result.Replace((char)8197), space); // 1/4 em-space (wide)
            //    result = result.Replace((char)160), space); // non-breaking space
            //}

            //// Unless we are preserving length, remove double-spaces
            //// This will also turn triple (and more) spaces into single spaces
            //// ReSharper disable once InvertIf
            //if (!preserveLength)
            //{
            //    int priorLength;
            //    do
            //    {
            //        priorLength = result.Length;
            //        result = result.Replace("  ", " ");
            //    } while (priorLength > result.Length);
            //}

            //return trimSpaces ? result.Trim() : result;
        }

        /// <summary>
        ///     Overload for Contains using the StringComparison options
        /// </summary>
        /// <param name="source">The source string</param>
        /// <param name="value">The string to seek.</param>
        /// <param name="comparisonType">One of the enumeration values that specifies the rules for the search.</param>

        public static bool Contains(this string source, string value, StringComparison comparisonType) =>
            source?.IndexOf(value, comparisonType) > -1;


        /// <summary>Determines the string contains any Smart Quote characters.See Also: IndexOfAny </summary>
        /// <param name="source">String containing sourceText to be queried for smart quotes.</param>
        /// <returns>Returns true if input sourceText contains any smart quotes.</returns>
        public static bool ContainsSmartQuotes(this string source) => source.Contains(SmartQuoteChars);


        /// <summary>Determines if the string contains any Plain Quote characters.</summary>
        public static bool ContainsPlainQuotes(this string source) => source.Contains(PlainQuoteChars);


        /// <summary>Determines if inputString contains any hyphens.</summary>
        /// <param name="source">String containing sourceText to be queried for hyphens.</param>
        /// <returns>Returns true if input sourceText contains any hyphens.</returns>
        public static bool ContainsHyphens(this string source) => source.Contains(HyphenChars);


        /// <summary>Determines if the string contains any of the characters in the characters argument. See Also: IndexOfAny</summary>
        /// <param name="source">String to be checked.</param>
        /// <param name="characters">An array of char containing the values for characters to be searched.</param>
        /// <returns>Returns true if the string contains any of the input characters.</returns>
        public static bool Contains(this string source, char[] characters) =>
            source != null &&
            source.IndexOfAny(characters) != -1; // //var chars = Array.ConvertAll(asciiValues, Convert.ToChar);


        /// <summary>Determines number of occurrences of a character within a string.</summary>
        /// <param name="source">String text to analyze.</param>
        /// <param name="c">Character to locate within the inputText.</param>
        /// <returns>Number of occurrences of character c within inputText.</returns>
        public static int CountOf(this string source, char c)
        {
#pragma warning disable CA1062 // Validate arguments of public methods
            return source.Split(c).Length - 1;
#pragma warning restore CA1062 // Validate arguments of public methods
        }

        /// <summary>Determines number of occurrences of a set of characters within a string.</summary>
        /// <param name="source">String text to analyze.</param>
        /// <param name="chars">Characters to locate within the inputText.</param>
        /// <returns>Number of occurrences of any of the characters in the input array within inputText.</returns>
        public static int CountOf(this string source, char[] chars)
        {
#pragma warning disable CA1062 // Validate arguments of public methods
            return source.Split(chars).Length - 1;
#pragma warning restore CA1062 // Validate arguments of public methods
        }

        /// <summary>Determines number of occurrences of a string within a string.</summary>
        /// <param name="source">String text to analyze.</param>
        /// <param name="text">String to locate within the inputText.</param>
        /// <returns>Number of occurrences of string text within inputText.</returns>
        public static int CountOf(this string source, string text)
        {
#pragma warning disable CA1062 // Validate arguments of public methods
            return source.Split(new[] { text }, StringSplitOptions.None).Length - 1;
#pragma warning restore CA1062 // Validate arguments of public methods
        }


        /// <summary></summary>
        /// <param name="source"></param>
        /// <param name="c"></param>
        /// <param name="count"></param>

        public static int IndexOf(this string source, char c, int count)
        {
            return IndexOfNthCore(source, c.ToString(), count);
        }

        /// <summary>Determines the index of the nth occurrence of a string within a string.</summary>
        /// <param name="source">String to analyze.</param>
        /// <param name="c">String to located within inputText.</param>
        /// <param name="count">Equal to the nth occurrence of c within inputText.</param>
        /// <returns>Index of the nth occurrence of c within inputText.</returns>
        public static int IndexOf(this string source, string c, int count)
        {
            return IndexOfNthCore(source, c, count);
        }

        private static int IndexOfNthCore(string inputText, string text, int count)
        {
            if (count <= 0 || string.IsNullOrEmpty(inputText) || string.IsNullOrEmpty(text))
                return -1;

            var idxChar = -1;
            var curCount = 0;
            while (curCount < count)
            {
                idxChar = inputText.IndexOf(text, idxChar + 1, StringComparison.Ordinal);
                if (idxChar == -1)
                {
                    return -1;
                }

                curCount += 1;
            }

            return idxChar;
        }


        /// <summary>Determines if first letter in input string is a capital letter.</summary>
        /// <param name="source">Text to analyze.</param>
        /// <returns>Returns true if first letter in string is between A and Z (inclusive).</returns>
        public static bool IsCapitalized(this string source)
        {


            if (string.IsNullOrEmpty(source))
            {
                return false;
            }

            source = source.Trim();

            return source[index: 0] >= 65 && source[index: 0] <= 90;
        }


        /// <summary>Determines if two strings are equal, using the current culture</summary>
        /// <param name="source">Source string</param>
        /// <param name="target">string for comparison</param>
        /// <param name="ignoreCase">Optionally ignore the case. Default value is true</param>

        public static bool IsEqual(this string source, string target, bool ignoreCase = true)
        {
            return string.Equals(source, target, ignoreCase ? StringComparison.CurrentCultureIgnoreCase : StringComparison.CurrentCulture);
        }


        /// <summary>Determines if two strings are identical, ignoring spaces and periods.</summary>
        /// <param name="source">Source text.</param>
        /// <param name="target">Text for comparison.</param>
        /// <param name="ignoreChars">An array of characters to ignore in the comparison.</param>
        /// <param name="ignoreCase"></param>

        public static bool IsEqual(this string source, string target, [CanBeNull] char[] ignoreChars, bool ignoreCase)
        {
            if (ignoreCase)
            {
                source = source?.ToLower();
                target = target?.ToLower();
            }

            if (ignoreChars == null)
            {
                return source == target;
            }

            foreach (var c in ignoreChars)
            {
                var charText = ignoreCase ? c.ToString().ToLower() : c.ToString();
                source = source?.Replace(charText, string.Empty);
                target = target?.Replace(charText, string.Empty);
            }


            return source == target;
        }


        /// <summary>
        ///     Returns the index of the item in the supplied string collection that matches the source string.
        /// </summary>
        /// <param name="source">The source string to compare.</param>
        /// <param name="items">A collection of strings to compare.</param>
        /// <param name="comparison">One of the enumeration values that specifies how the strings will be compared.</param>
        /// <returns>The index of the item in the collection that matched the source string; otherwise -1.</returns>
        public static int EqualsAny(this string source, [CanBeNull] IEnumerable<string> items, StringComparison comparison)
        {
            if (items == null)
                return -1;

            var counter = 0;
            foreach (var itm in items)
            {
                if (source.Equals(itm, comparison))
                    return counter;
                counter += 1;
            }

            return -1;
        }


        //Note using string as source here rather than char, as using char makes them harder to find.
        //AKA char.MaxValue.WhiteSpaceCharacters, vs string.Empty.WhiteSpaceChars;


        /// <summary>
        ///     Represents all the Unicode white space characters.
        ///     see https://en.wikipedia.org/wiki/Whitespace_character
        /// </summary>
        public static char[] WhiteSpaceCharacters(this string source) => WhiteSpaceChars;


        /// <summary>
        ///     Represents all the Unicode Hyphen characters.
        /// </summary>
        public static char[] HyphenCharacters(this string source) => HyphenChars;


        /// <summary>
        ///     Represents all the Unicode smart quote characters.
        /// </summary>
        public static char[] SmartQuoteCharacters(this string source) => SmartQuoteChars;


        /// <summary>
        ///     Represents all the Unicode plain quote characters.
        /// </summary>
        public static char[] PlainQuoteCharacters(this string source) => PlainQuoteChars;


        /// <summary>
        ///     Represents all the Unicode smart quote characters.
        /// </summary>
        public static char[] SpaceCharacters(this string source) => SpaceChars;

        /// <summary>
        ///     Represents no-width special characters (e.g., zero width non-joinger, zero width joinger, etc.)
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public static char[] NoWidthSpecialCharacters(this string source) => NoWidthSpecialChars;


        /// <summary>Returns if the string represents a Hyphen character.</summary>
        public static bool IsHyphen(this string source)
        {
            if (source == null || source.Length > 1)
                return false;

            return IsHyphen(source[index: 0]);
        }

        /// <summary>Returns if the string represents a Hyphen character.</summary>
        public static bool IsHyphen(this char source)
        {
            switch ((int)source)
            {
                case 0x002D:
                case 0x001E:
                case 0x001F:
                case 0x0096:
                case 0x0097:
                case 0x2010:
                case 0x2011:
                case 0x2012:
                case 0x2013:
                case 0x2014:
                case 0x2015:
                    return true;
            }

            return false;
        }

        /// <summary>Returns true if the string represents a no-width special character.</summary>
        public static bool IsNoWidthSpecialCharacter(this string source)
        {
            if (source == null || source.Length > 1)
                return false;
            return IsNoWidthSpecialCharacter(source[0]);
        }

        /// <summary>Returns true if the character is a no-width special character.</summary>
        public static bool IsNoWidthSpecialCharacter(this char source)
        {
            return NoWidthSpecialChars.Contains(source);
        }


        /// <summary>Returns if the string represents a smart quote character.</summary>
        public static bool IsSmartQuote(this string source)
        {
            if (source == null || source.Length > 1)
                return false;

            return IsSmartQuote(source[index: 0]);
        }

        /// <summary>Returns if the string represents a smart quote character.</summary>
        public static bool IsSmartQuote(this char source)
        {
            return SmartQuoteChars.Contains(source);
        }


        /// <summary>
        ///     Returns if the string starts with a plain quote character "(34) or '(39)
        /// </summary>
        /// <param name="source">The string to check.</param>
        public static bool IsPlainQuote(this string source)
        {
            if (source == null || source.Length > 1)
                return false;
            return IsPlainQuote(source[index: 0]);
        }

        /// <summary>
        ///     Returns if the character is a plain quote character "(34) or '(39)
        /// </summary>
        /// <param name="source">The string to check.</param>
        public static bool IsPlainQuote(this char source)
        {
            return PlainQuoteChars.Contains(source);
        }


        /// <summary>Returns if the string represents a Space character.</summary>
        public static bool IsSpace(this string source)
        {
            if (source == null || source.Length > 1)
                return false;

            return IsSpace(source[index: 0]);
        }

        /// <summary>Returns if the string represents a Space character.</summary>
        public static bool IsSpace(this char source)
        {
            switch ((int)source)
            {
                case 0x0020: //space
                case 0x00A0: //no-break space
                case 0x1680: //ogham space mark
                case 0x2000: //en quad
                case 0x2001: //em quad
                case 0x2002: //en space
                case 0x2003: //em space
                case 0x2004: //three-per-em space
                case 0x2005: //four-per-em space
                case 0x2006: //six-per-em space
                case 0x2007: //figure space
                case 0x2008: //punctuation space
                case 0x2009: //thin space
                case 0x200A: //hair space
                case 0x202F: //narrow no-break space
                case 0x205F: //medium mathematical space
                case 0x3000: //ideographic space
                    return true;
            }

            return false;
        }


        /// <summary>Determines if input character is a Word white-space character.</summary>
        /// <param name="source">The source character to check.</param>
        /// <returns>Returns true if input character is a white-space character.</returns>
        public static bool IsWhiteSpace(this char source)
        {
            switch ((int)source)
            {
                case 0x0009: //character tabulation
                case 0x000A: //line feed
                case 0x000B: //line tabulation
                case 0x000C: //form feed
                case 0x000D: //carriage return
                case 0x0020: //space
                case 0x0085: //next line
                case 0x00A0: //no-break space
                case 0x1680: //ogham space mark
                case 0x2000: //en quad
                case 0x2001: //em quad
                case 0x2002: //en space
                case 0x2003: //em space
                case 0x2004: //three-per-em space
                case 0x2005: //four-per-em space
                case 0x2006: //six-per-em space
                case 0x2007: //figure space
                case 0x2008: //punctuation space
                case 0x2009: //thin space
                case 0x200A: //hair space
                case 0x2028: //line separator
                case 0x2029: //paragraph separator
                case 0x202F: //narrow no-break space
                case 0x205F: //medium mathematical space
                case 0x3000: //ideographic space
                    return true;
            }

            return false;
        }

        /// <summary>Determines if input string contains only Word white-space characters.</summary>
        /// <param name="source">Input string to check.</param>
        /// <returns>Returns true if all characters in the input sourceText are Word white-space characters.</returns>
        public static bool IsWhiteSpace(this string source)
        {
            return !string.IsNullOrEmpty(source) && string.IsNullOrWhiteSpace(source);
        }


        /// <summary>
        ///     Removes all the white space characters from the supplied string.
        /// </summary>
        /// <param name="source">The string to remove the white space characters from</param>
        /// <returns>A new string with all the white space characters removed</returns>
        public static string RemoveWhitespace(this string source)
        {
            return new string(source.ToCharArray()
                                    .Where(c => !char.IsWhiteSpace(c))
                                    .ToArray());
        }


        /// <summary>Removes text from start of text string.</summary>
        /// <param name="sourceText">Input text to be trimmed.</param>
        /// <param name="itemsToTrim">Separated list of strings to removed from the start of the input text.</param>
        /// <param name="ignoreCase">If true, will not trim text if case does not match.</param>
        /// <param name="trimWhiteSpace">If true, will automatically trim all whitespace from input text.</param>

        public static string TrimStart(this string sourceText, bool ignoreCase, bool trimWhiteSpace, params string[] itemsToTrim)
        {
            return TrimCore(sourceText, trimEnd: false, ignoreCase: ignoreCase, trimWhiteSpace: trimWhiteSpace, itemsToTrim: itemsToTrim);
        }


        /// <summary>Removes text from end of text string.</summary>
        /// <param name="source">Input text to be trimmed.</param>
        /// <param name="itemsToTrim">A list of strings to removed from the end of the input text.</param>
        /// <param name="ignoreCase">If true, will not trim text if case does not match.</param>
        /// <param name="trimWhiteSpace">If true, will automatically trim all whitespace from input text.</param>

        public static string TrimEnd(this string source, bool ignoreCase, bool trimWhiteSpace, params string[] itemsToTrim)
        {
            return TrimCore(source, trimEnd: true, ignoreCase: ignoreCase, trimWhiteSpace: trimWhiteSpace, itemsToTrim: itemsToTrim);
        }


        /// <summary>Removes text from end of text string.</summary>
        /// <param name="source">Input text to be trimmed.</param>
        /// <param name="itemsToTrim">Separated list of strings to removed from the end of the input text.</param>
        /// <param name="ignoreCase">If true, will not trim text if case does not match.</param>
        /// <param name="trimWhiteSpace">If true, will automatically trim all whitespace from input text.</param>

        public static string Trim(this string source, bool ignoreCase, bool trimWhiteSpace, string[] itemsToTrim)
        {
            var result = TrimStart(source, ignoreCase, trimWhiteSpace, itemsToTrim);
            return TrimEnd(result, ignoreCase, trimWhiteSpace, itemsToTrim);
        }


        private static string TrimCore(this string sourceText, bool trimEnd, bool ignoreCase, bool trimWhiteSpace, params string[] itemsToTrim)
        {
            if (string.IsNullOrEmpty(sourceText))
            {
                return sourceText;
            }

            foreach (var textToTrim in itemsToTrim)
            {
                var comparer = ignoreCase ? StringComparison.CurrentCultureIgnoreCase : StringComparison.CurrentCulture;
                if (trimEnd)
                    while (sourceText.EndsWith(textToTrim, comparer))
                        sourceText = sourceText.Substring(0, sourceText.Length - textToTrim.Length);
                else
                    while (sourceText.StartsWith(textToTrim, comparer))
                        sourceText = sourceText.Substring(textToTrim.Length);
            }

            if (trimWhiteSpace)
                sourceText = trimEnd ? sourceText.TrimEnd() : sourceText.TrimStart();

            return sourceText;
        }


        /// <summary>Returns text in reverse.</summary>
        public static string Reverse(this string source)
        {
            return new string(source.ToArray().Reverse().ToArray());
        }


        /// <summary>
        ///     Truncates and adds "..." to the end of the string, if the string is longer than the maxChars value
        /// </summary>
        /// <param name="source">The string to truncate</param>
        /// <param name="maxChars">The maximum number of characters allowed before truncating the string</param>

        public static string Truncate(this string source, int maxChars)
        {
            if (source == null)
                return null;
            return source.Length <= maxChars ? source : source.Substring(0, maxChars) + "...";
        }


        /// <summary>
        ///     Parses the string into a double value. If the string is null or cannot be passed zero is returned
        /// </summary>
        /// <param name="source">The string to parse</param>
        /// <returns>A double</returns>
        public static double Val([CanBeNull] this string source)
        {
            if (source == null)
                return 0;

            try
            {
                //try the entire string, then progressively smaller
                //substrings to simulate the behavior of VB's 'Val',
                //which ignores trailing characters after a recognizable value:
                for (var size = source.Length; size > 0; size--)
                {
                    if (double.TryParse(source.Substring(0, size), out var testDouble))
                        return testDouble;
                }
            }
            catch
            {
                //Ignore
            }

            //no value is recognized, so return 0:
            return 0;
        }


        /// <summary>Core routine that modifies key to ensure that it is unique. Typically used for file names.</summary>
        /// <param name="source">An ICollection if items containing possible keys that are duplicates.</param>
        /// <param name="startKey">Item to de-duplicate.</param>
        /// <param name="format">Format string to rename when de-duplicating.</param>
        /// <returns>Returns new key value.</returns>
        public static string MakeUniqueKey([NotNull] this IEnumerable<string> source, [NotNull] string startKey, [NotNull] string format = "{0} - Copy ({1})")
        {
            Check.NotNull(source, nameof(source));
            Check.NotEmpty(startKey, nameof(startKey));
            Check.NotEmpty(format, nameof(format));

            var nKey = 1;

            var newKey = startKey;
            do
            {
                if (!source.Contains(newKey))
                    return newKey;

                newKey = string.Format(CultureInfo.CurrentCulture, format, startKey, nKey);
                nKey += 1;
            } while (true);
        }


        /// <summary>
        ///     Inserts returns in text at or before maxLineLength only where standard spaces exist,
        ///     after performing a Clean on spaces, hyphens, and quotes. Used for text output.
        /// </summary>
        /// <param name="source">Text to be cleaned and wrapped.</param>
        /// <param name="maxLineLength">Maximum number of characters per line.</param>
        /// <returns>Returns cleaned text with line breaks inserted.</returns>
        public static string WrapClean(this string source, int maxLineLength)
        {
            source = source.Clean();
            return source.Wrap(maxLineLength);
        }

        /// <summary>Inserts returns in text at or before maxLineLength only where standard spaces exist.</summary>
        /// <param name="source">Text to be wrapped.</param>
        /// <param name="maxLineLength">Maximum number of characters per line.</param>
        /// <returns>Returns text with line breaks inserted.</returns>
        public static string Wrap(this string source, int maxLineLength)
        {
            const char spaceChar = ' ';

            var words = source.Split(new[] { spaceChar }, StringSplitOptions.RemoveEmptyEntries);
            var sb = new StringBuilder();
            var line = string.Empty;
            foreach (var word in words)
            {
                var newLineLength = line.Length + word.Length + 1; // include space char in calculation
                if (line.Length != 0 && newLineLength > maxLineLength)
                {
                    sb.AppendLine(line.TrimEnd(spaceChar));
                    line = string.Empty;
                }

                line += $"{word} ";
            }

            // Include last (partial) line if it exists
            if (line.Length > 0)
                sb.AppendLine(line.TrimEnd(spaceChar));

            return sb.ToString();
        }


        /// <summary>Splits the string into chunks with input length.</summary>
        /// <param name="s">The string to split.</param>
        /// <param name="length">The size of each chunk.</param>
        /// <returns>An IEnumerable&gt;string&lt; of chunks of type string.</returns>
        public static IEnumerable<string> SplitByLength(this string s, int length)
        {
            for (var i = 0; i < s.Length; i += length)
            {
                if (i + length <= s.Length)
                    yield return s.Substring(i, length);
                else
                    yield return s.Substring(i);
            }
        }

        /// <summary>Gets the text in a string before the first occurrence of an input character.</summary>
        /// <param name="s">The source string.</param>
        /// <param name="c">The character to search for.</param>
        /// <returns>Returns the text in s that exists before the first occurrence of c.</returns>
        public static string TextBeforeChar(this string s, char c)
        {
            if (!s.Contains(c))
                return s;

            var idx = s.IndexOf(c);
            return s.Substring(0, idx);
        }

        /// <summary>Gets the text in a string before the first occurrence of any characters from an input set.</summary>
        /// <param name="s">The source string.</param>
        /// <param name="chars">The characters to search for.</param>
        /// <returns>Returns the text in s that exists before the first occurrence of any characters in chars.</returns>
        public static string TextBeforeChars(this string s, IEnumerable<char> chars)
        {
            var idxMin = 9999;
            foreach (var c in chars)
            {
                var idx = s.IndexOf(c);
                if (idx >= 0)
                    idxMin = Math.Min(idxMin, idx);
            }

            return s.Substring(0, idxMin);
        }

        /// <summary>Gets the text in a string after the last occurrence of an input character.</summary>
        /// <param name="s">The source string.</param>
        /// <param name="c">The character to search for.</param>
        /// <returns>Returns the text in s that exists after the last occurrence of c.</returns>
        public static string TextAfterChar(this string s, char c)
        {
            if (!s.Contains(c))
                return null;

            var idx = s.LastIndexOf(c);
            return s.Substring(idx + 1);
        }

        /// <summary>Converts a string to a delimited string of ASCII values</summary>
        /// <param name="s">The source string</param>
        /// <param name="separator">Character used to separate the ASCII values. Default is a pipe ('|')</param>
        /// <returns>Returns the delimited text.</returns>
        public static string ToDelimitedAscii(this string s, char separator = '|')
        {
            var retVal = string.Empty;
            foreach (var chr in s.ToCharArray())
            {
                retVal += ((int)chr).ToString();
                retVal += separator.ToString();
            }

            return retVal.TrimEnd(separator);
        }

        /// <summary>
        /// Pluralizes the source string if the amount is greater than one.
        /// </summary>
        /// <param name="source">The string to pluralize. the string should contain a single word.</param>
        /// <param name="amount"></param>
        /// <param name="cultureInfo">if null the CurrentUICulture is used.</param>
        /// <returns>The source string or a pluralized version of the string</returns>
        public static string Pluralize(this string source, int amount, CultureInfo cultureInfo = null)
        {

            if (amount == 1 || string.IsNullOrEmpty(source))
                return source;
 
            if (cultureInfo == null)
                cultureInfo = CultureInfo.CurrentUICulture;
            return PluralizationService.CreateService(cultureInfo).Pluralize(source);
        }
    }
}