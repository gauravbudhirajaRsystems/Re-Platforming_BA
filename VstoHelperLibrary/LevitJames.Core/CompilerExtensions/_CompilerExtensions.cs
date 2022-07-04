// © Copyright 2018 Levit & James, Inc.

using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using JetBrains.Annotations;

namespace LevitJames.Core
{
    /// <summary>Provides CompilerExtensions for existing framework types.</summary>

    //[DebuggerStepThrough]
    //[EditorBrowsable(EditorBrowsableState.Never)]
    public static partial class CompilerExtensions
    {
        /// <summary>
        ///     Returns if the type passed is a numeric type.
        ///     <para>The Numeric types checked for are: byte, short, int, long, float, double, decimal, sbyte, ushort, uint, ulong</para>
        /// </summary>
        /// <param name="source">A Type instance.</param>
        public static bool IsNumericType(this Type source)
        {
            return IsNumericType(source, out _);
        }

        /// <summary>Returns is the type passed is a numeric type.</summary>
        /// <overloads>
        ///     <para>
        ///         The Numeric types that are checked for are: byte,short,int,long,float,double,decimal,ushort,uint,ulong and
        ///         sbyte.
        ///     </para>
        ///     <para></para>
        /// </overloads>
        /// <param name="source">A Type instance.</param>
        /// <param name="isFixed">Returns true if the numeric type is fixed or a floating numeric type.</param>
        public static bool IsNumericType(this Type source, out bool isFixed)
        {
            switch (Type.GetTypeCode(source))
            {
            case TypeCode.Byte:
            case TypeCode.SByte:
            case TypeCode.UInt16:
            case TypeCode.UInt32:
            case TypeCode.UInt64:
            case TypeCode.Int16:
            case TypeCode.Int32:
            case TypeCode.Int64:
                isFixed = true;
                return true;
            case TypeCode.Decimal:
            case TypeCode.Double:
            case TypeCode.Single:
                isFixed = false;
                return true;
            default:
                isFixed = false;
                return false;
            }
        }


        /// <summary>Returns the matching bracket character, or null if the input character is not a bracket character.</summary>
        /// <param name="c"></param>
        /// <returns>The supported bracket characters are &lt;{([])}&gt;</returns>
        public static char MatchingBracket(this char c)
        {
            var position = @"<{([])}>".IndexOf(c);
            return @">})][({<"[position];
        }


        /// <summary>Gets maximum value from an array of integers.</summary>
        /// <param name="values">Integer array to analyze.</param>
        /// <returns>Maximum value in array.</returns>
        public static int IntMax(this int[] values)
        {
            return values.Concat(new[] {int.MinValue}).Max();
        }


        /// <summary>Gets minimum value from an array of integers.</summary>
        /// <param name="values">Integer array to analyze.</param>
        /// <returns>Minimum value in array.</returns>
        public static int IntMin(this int[] values)
        {
            return values.Concat(new[] {int.MaxValue}).Min();
        }


        /// <summary>
        ///     Determines if date text is in mm/dd/yy or mm/dd/yyyy format,
        ///     with leading 0s optional for mm &amp; dd. Also accepts dd/mm/yy or dd/mm/yyyy.
        /// </summary>
        /// <param name="dateText">Date text to analyze.</param>
        /// <returns>Returns true if date text is in slashed date format.</returns>
        public static bool IsSlashedDate(this string dateText)
        {
            var tokens = dateText.Split('/');
            var upperLimit = tokens.GetUpperBound(0);

            if (upperLimit < 1 || upperLimit > 2)
            {
                return false;
            }

            for (var position = 0; position <= tokens.GetUpperBound(0); position++)
            {
                var token = tokens[position];

                // Validate the length
                var length = token.Length;
                if (length == 1 && position < upperLimit)
                {
                    // ok
                }
                else if (length == 2)
                {
                    // ok
                }
                else if (length == 4 && position == upperLimit)
                {
                    // ok
                }
                else
                {
                    return false;
                }

                // Get the value
                if (!int.TryParse(token, out var value))
                {
                    value = 0;
                }

                if (value < 1)
                {
                    return false;
                }

                if (position < upperLimit && value > 31)
                {
                    return false;
                }

                // Convert back to string, and see if they match.
                // This will filter out non-integers, non-pure numbers etc.
                var valueString = value.ToString().Trim();
                if (token == valueString)
                {
                    // ok
                }
                else if (length == 2 && (token == "0" + valueString))
                {
                    // ok e.g. "07" matches value 7
                }
                else
                {
                    return false;
                }
            }

            return true;
        }


        /// <summary>
        ///     Appends the string value to the StringBuilder class using the provided indent level followed by the default line
        ///     terminator.
        /// </summary>
        /// <param name="source">A valid StringBuilder instance to append the string to.</param>
        /// <param name="value">The string to append</param>
        /// <param name="indentLevel">The number of spaces to append to string.</param>
        public static void AppendIndentedLine(this StringBuilder source, string value, int indentLevel)
        {
            if (indentLevel > 0)
                source.Append(new string(' ', indentLevel));

            source.AppendLine(value);
        }


        /// <summary>
        ///     Appends the string value to the StringBuilder class using the provided indent level.
        /// </summary>
        /// <param name="source">A valid StringBuilder instance to append the string to.</param>
        /// <param name="value">The string to append</param>
        /// <param name="indentLevel">The number of spaces to append to string.</param>
        public static void AppendIndented(this StringBuilder source, string value, int indentLevel)
        {
            if (indentLevel > 0)
                source.Append(new string(' ', indentLevel));

            source.Append(value);
        }

 
        /// <summary>
        /// Compares the contents of two file streams to see if they are the same.
        /// </summary>
        /// <param name="source">The source file</param>
        /// <param name="compare">A file to compare to the source file.</param>
        /// <returns>true if the file contents are the same; false otherwise</returns>
        public static bool ContentsEqual([NotNull] this FileInfo source, FileInfo compare)
        {
            return ContentsEqual(source, compare, Equals);
        }

        private static bool ContentsEqual([NotNull] this FileInfo source, FileInfo compare, Func<Stream, Stream, bool> areStreamsIdentical)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            if (compare == null)
                throw new ArgumentNullException(nameof(compare));

            if (areStreamsIdentical == null)
                throw new ArgumentNullException(nameof(areStreamsIdentical));

            if (!source.Exists || !compare.Exists)
                return false;

            if (source.Length != compare.Length)
                return false;

            using (var thisFile = source.OpenRead())
            using (var valueFile = compare.OpenRead())
            {
                if (valueFile.Length != thisFile.Length)
                    return false;

                if (!areStreamsIdentical(thisFile, valueFile))
                    return false;
            }
            return true;
        }

        /// <summary>
        /// Compares two stream instances to see if they are the same.
        /// </summary>
        /// <param name="source">The source stream</param>
        /// <param name="compare">The stream to compare with the source stream</param>
        /// <returns>true if the stream contents are the same; false otherwise</returns>
        public static bool Equals([NotNull] this Stream source, [NotNull] Stream compare)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            if (compare == null)
                throw new ArgumentNullException(nameof(compare));

            if (ReferenceEquals(source, compare))
                return true;

            const int buferSize = 80000; // 80000 is below LOH (85000)
            var buffer1 = new byte[buferSize];
            var buffer2 = new byte[buferSize];
 
            do
            {
                int read1 = source.Read(buffer1, 0, buffer1.Length);
                if (read1 == 0)
                    return compare.Read(buffer2, 0, 1) == 0; // check not eof

                // both stream read could return different counts
                int read2 = 0;
                do
                {
                    int read3 = compare.Read(buffer2, read2, read1 - read2);
                    if (read3 == 0)
                        return false;

                    read2 += read3;
                }
                while (read2 < read1);

                if (!Equals(buffer1, buffer2))
                    return false;
            }
            while (true);
        }

        /// <summary>
        /// Compares two byte arrays to see if they hold the same contents.
        /// </summary>
        /// <param name="source">The source array</param>
        /// <param name="compare">The byte array to compare with the source array</param>
        /// <returns>true if the byte array contents are the same; false otherwise</returns>
        public static bool Equals([NotNull] this byte[] source, byte[] compare)
        {

            if (source == null)
                throw new ArgumentNullException(nameof(source));

            if (compare == null)
                throw new ArgumentNullException(nameof(compare));

            if (source.Length != compare.Length)
                return false;

            for (var i = 0; i < source.Length; i++)
            {
                if (source[i] != compare[i])
                    return false;
            }
            return true;
        }

        /// <summary>Returns the standard text string for an assembly version.</summary>
        /// <param name="source">The calling assembly.</param>
        /// <returns>Returns the version text of the calling assembly.</returns>
        // ReSharper disable once InconsistentNaming
        public static string VersionTextLJ([NotNull] this Assembly source)
        {
            return source.GetName().Version.ToString();
        }

        /// <summary>Returns the text string for an assembly version using the specified formatting.</summary>
        /// <param name="source">The calling assembly.</param>
        /// <param name="formatText">The format string for the version text.</param>
        /// <returns>Returns the version text of the calling assembly.</returns>
        // ReSharper disable once InconsistentNaming
        public static string VersionTextLJ([NotNull] this Assembly source, string formatText)
        {
            return string.Format(formatText, source.GetName().Version);
        }


        /// <summary>
        ///     Compares two float values and determines if they are the same value within the tolerance provided
        /// </summary>
        /// <param name="source">the source value</param>
        /// <param name="compareTo">The value to compare the source value with</param>
        /// <param name="tolerance">The tolerance within which the values are deemed equal. Default is 0.001</param>
        /// <returns>
        ///     One if the source is greater than the compareTo value.
        ///     -1 if the source is less than the compareTo value; zero if the values are deemed equal.
        /// </returns>
        public static bool AreClose(this float source, float compareTo, float tolerance = 0.001F)
        {
            if (source - compareTo > tolerance)
                return false;
            if (compareTo - source > tolerance)
                return false;
            return true;
        }

        /// <summary>
        ///     Compares two double values and determines if they are the same value within the tolerance provided
        /// </summary>
        /// <param name="source">the source value</param>
        /// <param name="compareTo">The value to compare the source value with</param>
        /// <param name="tolerance">The tolerance within which the values are deemed equal. Default is 0.001</param>
        /// <returns>
        ///     One if the source is greater than the compareTo value.
        ///     -1 if the source is less than the compareTo value; zero if the values are deemed equal.
        /// </returns>
        public static bool AreClose(this double source, double compareTo, double tolerance = double.Epsilon)
        {
            if (source - compareTo > tolerance)
                return false;
            if (compareTo - source > tolerance)
                return false;
            return true;
        }
    }
}