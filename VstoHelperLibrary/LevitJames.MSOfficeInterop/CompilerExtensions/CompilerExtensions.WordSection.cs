// © Copyright 2020 Levit & James, Inc.

using JetBrains.Annotations;
using LevitJames.Core;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace LevitJames.MSOffice.MSWord
{
    public static partial class Extensions
    {
        /// <summary>Determines if any header in the source has any content.</summary>
        /// <param name="source">Word source to analyze.</param>
        /// <returns>Returns true, if there is a non-empty header in the source.</returns>
        public static bool AnyHeadersWithContent([NotNull] this Section source)
        {
            CheckNotNull(source);
            var headers = source.Headers;
            var result = false;

            for (var idx = WdHeaderFooterIndex.wdHeaderFooterPrimary; idx <= WdHeaderFooterIndex.wdHeaderFooterEvenPages; idx++)
            {
                var header = headers[idx];
                if (!header.Exists)
                {
                    Marshal.ReleaseComObject(header);
                    continue;
                }

                var headerRange = header.Range;
                Marshal.ReleaseComObject(header);

                if (headerRange.End - headerRange.Start == 1)
                    result = headerRange.Text != "\n";

                Marshal.ReleaseComObject(headerRange);
                break;
            }

            Marshal.ReleaseComObject(headers);

            return result;
        }


        /// <summary>Determines if any footer in the source has any content.</summary>
        /// <param name="source">Word source to analyze.</param>
        /// <returns>Returns true, if there is a non-empty footer in the source.</returns>
        public static bool AnyFootersWithContent([NotNull] this Section source)
        {
            CheckNotNull(source);

            var footers = source.Footers;
            var result = false;

            for (var idx = WdHeaderFooterIndex.wdHeaderFooterPrimary; idx <= WdHeaderFooterIndex.wdHeaderFooterEvenPages; idx++)
            {
                var footer = footers[idx];
                if (!footer.Exists)
                {
                    Marshal.ReleaseComObject(footer);
                    continue;
                }

                var footerRange = footer.Range;
                Marshal.ReleaseComObject(footer);

                if (footerRange.End - footerRange.Start == 1)
                    result = footerRange.Text != "\n";
  
                Marshal.ReleaseComObject(footerRange);
                break;
            }

            Marshal.ReleaseComObject(footers);

            return result;
        }


        /// <summary>
        ///     Returns the Range for which a Section's properties are effective.
        ///     If there are preceding Sections with deleted section marks, the
        ///     range of those sections is include as well
        /// </summary>
        /// <param name="source">A Section instance.</param>
        /// <returns>A Word.Range.</returns>
        // ReSharper disable once InconsistentNaming
        public static Range RangeLJ([NotNull] this Section source)
        {
            Check.NotNull(source, nameof(source));

            var rng = source.Range;

            var docSecs = ((Document) source.Parent).Sections;

            for (var idx = source.Index - 1; idx >= 1; idx--)
            {
                var sec = docSecs[idx];
                if (sec.SectionBreakHasBeenDeletedLJ())
                {
                    var rng2 = sec.Range;
                    rng.Start = rng2.Start;
                    Marshal.ReleaseComObject(rng2);
                    Marshal.ReleaseComObject(sec);
                }
                else
                {
                    Marshal.ReleaseComObject(sec);
                    break;
                }
            }

            Marshal.ReleaseComObject(docSecs);

            return rng;
        }


        /// <summary>
        ///     Returns the Range for the end-of-section character
        ///     For the final Section in a document, this appears as a paragraph mark
        /// </summary>
        /// <param name="source">A Section instance.</param>
        /// <returns>A Word.Range object.</returns>
        // ReSharper disable once InconsistentNaming
        public static bool SectionBreakHasBeenDeletedLJ([NotNull] this Section source)
        {
            Check.NotNull(source, nameof(source));


            // There is no section break for the last section
            var range = source.RangeLJ();

            if (source.Index == range.Document.Sections.Count)
                return false;

            var sectionMarkRange = source.SectionMarkRangeLJ();
            if (sectionMarkRange == null)
                return false;

            // KDP NOTES 2019/12/23
            // TSWA 7241: There is a Word bug related to revisions in the same Word range
            // The Revisions property returns the effective revisions for that range, but
            // you must access the revisions by index. If you use the enumerator, you will
            // get all of the revisions for that range. Consequently, using the RevisionsLJ()
            // compiler extension does not work for this either.
            //
            // KDP Note: There is no property of the Revision that I can see that can tell
            // you whether the revision has been superceded.
            var deleted = false;
            for (var idx = 1; idx <= sectionMarkRange.Revisions.Count; idx++)
            {
                var rev = sectionMarkRange.Revisions[idx];
                deleted = rev.Type == WdRevisionType.wdRevisionDelete;
                Marshal.ReleaseComObject(rev);
                if (deleted)
                    break;
            }

            // using ReleaseComObject on either sectionMarkRange or revisions also release the other 
            Marshal.ReleaseComObject(sectionMarkRange);
 
            return deleted;
        }


        /// <summary>
        ///     Returns the Range for the end-of-section character
        ///     For the final Section in a document, this appears as a paragraph mark
        /// </summary>
        /// <param name="source">A Section instance.</param>
        /// <returns>A Word.Range object.</returns>
        // ReSharper disable once InconsistentNaming
        public static Range SectionMarkRangeLJ([NotNull] this Section source)
        {
            Check.NotNull(source, nameof(source));

            var rng = source.Range;
            rng.Start = rng.End - 1;
            return rng;
        }


        /// <summary>Gets the width of the section text, in points</summary>
        /// <param name="source">A Section instance.</param>
        /// <returns>
        ///     The width of the text (in points). If any of the PageWidth,
        ///     LeftMargin, RightMargin, or Gutter values are undefined, returns 9999990F.
        /// </returns>
        // ReSharper disable once InconsistentNaming
        public static float TextWidthLJ([NotNull] this Section source)
        {
            var pageSetup = source.PageSetup;

            var pw = pageSetup.PageWidth;
            var lm = pageSetup.LeftMargin;
            var rm = pageSetup.RightMargin;
            var gu = pageSetup.Gutter;
            Marshal.ReleaseComObject(pageSetup);

            // If any of the values is undefined, return an "undefined value".
            const float minUndefined = (float) WdConstants.wdUndefined;
            if (pw > minUndefined || lm > minUndefined || rm > minUndefined || gu > minUndefined)
                return minUndefined;

            // All values are proper; return calculated value.
            return pw - lm - rm - gu;
        }
    }
}