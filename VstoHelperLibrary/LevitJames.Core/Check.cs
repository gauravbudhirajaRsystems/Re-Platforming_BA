// © Copyright 2018 Levit & James, Inc.

using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
namespace LevitJames.Core
{
    /// <summary>
    ///     Used by the VS CodeAnalisys tool to filter out CA1062 warnings when the null reference is checked in a helper
    ///     method
    /// </summary>
    [AttributeUsage(AttributeTargets.Parameter, AllowMultiple = true)]
    internal sealed class ValidatedNotNullAttribute : Attribute { }


    /// <summary>
    ///     A class for checking input parameters and throwing an exception if the criteria is not met.
    /// </summary>
    [DebuggerStepThrough]
    public static class Check
    {
        /// <summary>
        ///     Checks if the value passed is not null.
        /// </summary>
        /// <typeparam name="T">The type of value to check.</typeparam>
        /// <param name="value">The value to check.</param>
        /// <param name="parameterName">The name of the parameter to check.</param>
        /// <exception cref="ArgumentNullException">Throws an ArgumentNullException if the value is null.</exception>
        public static void NotNull<T>([ValidatedNotNull] T value, string parameterName) where T : class
        {
            if (value == null)
            {
                throw new ArgumentNullException(parameterName);
            }
        }

        /// <summary>
        ///     Checks if the value passed is not null.
        /// </summary>
        /// <typeparam name="T">The type of value to check.</typeparam>
        /// <param name="value">The value to check.</param>
        /// <param name="parameterName">The name of the parameter to check.</param>
        /// <exception cref="ArgumentNullException">Throws an ArgumentNullException if the value is null.</exception>
        public static void NotNull<T>([ValidatedNotNull] T? value, string parameterName) where T : struct
        {
            if (!value.HasValue)
            {
                throw new ArgumentNullException(parameterName);
            }
        }

        /// <summary>
        ///     Checks whether the Enum value passed is a correctly value.
        /// </summary>
        /// <param name="value">The value of the Enum.</param>
        /// <param name="parameterName">The name of the parameter to check.</param>
        /// <exception cref="ArgumentNullException">Throws an ArgumentNullException if the value is null.</exception>
        public static void Enum([ValidatedNotNull] Enum value, string parameterName)
        {
            if (!System.Enum.IsDefined(value.GetType(), value))
            {
                throw new ArgumentNullException(parameterName);
            }
        }

        /// <summary>
        ///     Checks if a Flagged Enum value passed matches the mask passed.
        /// </summary>
        /// <param name="value">The value of the Enum.</param>
        /// <param name="mask">A mask of Enum values to check against.</param>
        /// <param name="parameterName">The name of the parameter to check.</param>
        /// <exception cref="ArgumentOutOfRangeException">Throws an ArgumentOutOfRangeException if the value is null.</exception>
        public static void Enum(Enum value, Enum mask, string parameterName)
        {
            NotNull(value, parameterName);

            var intValue = Convert.ToInt32(value, CultureInfo.InvariantCulture);
            var intMask = Convert.ToInt32(mask, CultureInfo.InvariantCulture);

            if ((intValue & intMask) != intMask)
            {
                throw new ArgumentOutOfRangeException(parameterName);
            }
        }

        /// <summary>
        ///     Checks if the string value passed is not null or contains just white space.
        /// </summary>
        /// <param name="value">The string value to check.</param>
        /// <param name="parameterName">The name of the parameter to check.</param>
        public static void NotEmpty([ValidatedNotNull] string value, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new ArgumentNullException(parameterName, "The string passed is null or contains just white space.");
            }
        }


        /// <summary>
        ///     Checks if the string value passed meets the passed length requirements.
        /// </summary>
        /// <param name="value">The value to check.</param>
        /// <param name="parameterName">The name of the parameter to check.</param>
        /// <param name="min">The minimum required length of the string. Can be null if there is no maximum length.</param>
        /// <param name="max">The maximum required length of the string. Can be null if there is no maximum length.</param>
        /// <exception cref="ArgumentOutOfRangeException">
        ///     Throws an ArgumentNullException if the value does not meet either the min
        ///     or max length required.
        /// </exception>
        public static void Length(string value, string parameterName, int? min = null, int? max = null)
        {
            if (min.HasValue && value?.Length < min)
                throw new ArgumentOutOfRangeException(parameterName, parameterName + "is less than " + min);

            if (max.HasValue && value.Length > max)
                throw new ArgumentOutOfRangeException(parameterName, parameterName + "is greater than " + min);
        }

        /// <summary>Checks if the file path passed exists.</summary>
        /// <param name="filePath">The file path to check.</param>
        /// <param name="parameterName">The name of the parameter to check.</param>
        /// <exception cref="ArgumentException">
        ///     Throws an ArgumentException if the file path does not exist.
        /// </exception>
        public static void FileExists([ValidatedNotNull] string filePath, string parameterName)
        {
            if (!File.Exists(filePath))
            {
                throw new ArgumentException(parameterName, $"File: {parameterName} does not exist.");
            }
        }
    }
}
