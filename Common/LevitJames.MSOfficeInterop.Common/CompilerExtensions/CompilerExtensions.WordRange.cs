using LevitJames.MSOfficeInterop.Common.Internal;
using LevitJames.Shared.Common;
using Microsoft.Office.Interop.Word;

namespace LevitJames.MSOfficeInterop.Common.CompilerExtensions
{
    public static partial class Extensions
    {
        /// <summary>
        ///     Returns a Word.Range from a Word.Rectangle.
        ///     Unlike the Word supplied Rectangle.Range member this method will not throw an exception if the Range does not
        ///     exist.
        /// </summary>
        /// <param name="source">A Range instance.</param>
        /// <returns>A Word.Range, or Null if the Word.Range does not exist</returns>

        // ReSharper disable once InconsistentNaming
        public static Range RangeLJ([NotNull] this Rectangle source)
        {
            Check.NotNull(source, nameof(source));
            // ReSharper disable once SuspiciousTypeConversion.Global
            var hr = ((Rectangle12)source).Range(out var range);
            return hr == 0 ? range : null;
        }
    }
}
