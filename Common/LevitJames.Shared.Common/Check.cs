using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;

namespace LevitJames.Shared.Common
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
    }
}
