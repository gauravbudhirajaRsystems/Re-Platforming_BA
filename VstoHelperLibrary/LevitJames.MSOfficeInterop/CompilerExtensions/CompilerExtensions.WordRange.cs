// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using JetBrains.Annotations;
using LevitJames.Core;
using LevitJames.MSOffice.Internal;
using LevitJames.TextServices;
using Microsoft.Office.Interop.Word;

namespace LevitJames.MSOffice.MSWord
{
    public static partial class Extensions
    {
        /// <summary>
        ///     Returns a list of Sections which overlap a Range,
        ///     but excluding the sections associated with a deleted section break
        /// </summary>
        /// <param name="source">A Range instance.</param>
        /// <returns>A list of Word.Revision objects.</returns>
        // ReSharper disable once InconsistentNaming
        public static IEnumerable<Section> LJSectionsExceptDeleted([NotNull] this Range source)
        {
            Check.NotNull(source, nameof(source));

            Section sec = null;
            var secWasDeleted = false;

            foreach (Section secWithinLoop in source.Sections)
            {
                sec = secWithinLoop;
                secWasDeleted = secWithinLoop.SectionBreakHasBeenDeletedLJ();
                if (!secWasDeleted)
                    yield return secWithinLoop;
                else
                    Marshal.ReleaseComObject(secWithinLoop);
            }

            if (!secWasDeleted)
                yield break;

            var docSecs = ((Document)sec.Parent).Sections;
            while (sec.SectionBreakHasBeenDeletedLJ() && sec.Index < docSecs.Count)
            {
                var secIndex = sec.Index + 1;
                Marshal.ReleaseComObject(sec);
                sec = docSecs[secIndex];
            }

            Marshal.ReleaseComObject(docSecs);

            // I don't think it's possible for final section break to be deleted,
            // but even so, we will need to use it.
            yield return sec;
        }

        /// <summary>
        ///     Returns a list of Sections which overlap a Range,
        ///     but excluding the sections associated with a deleted section break
        /// </summary>
        /// <param name="source">A Range instance.</param>

        // ReSharper disable once InconsistentNaming
        public static IEnumerable<Section> LJSectionsExceptDeleted([NotNull] this Document source)
        {
            return source.Content.LJSectionsExceptDeleted();
        }


        /// <summary>
        ///     Returns the section pointer whose properties are effective for the given input Range.
        /// </summary>
        /// <param name="source">A Range instance.</param>
        /// <returns>A Word.Section object.</returns>
        /// <remarks>
        ///     If the input Range spans multiple non-deleted sections, the first
        ///     non-deleted section will be returned.
        /// </remarks>
        // ReSharper disable once InconsistentNaming
        public static Section EffectiveSectionLJ([NotNull] this Range source)
        {
            Check.NotNull(source, nameof(source));
            var sections = source.Sections;
            var firstSection = sections.First;
            var firstSectionRange = firstSection.RangeLJ();
            var firstSectionRangeNotDeleted = firstSectionRange?.LJSectionsExceptDeleted().FirstOrDefault();

            Marshal.ReleaseComObject(sections);
            Marshal.ReleaseComObject(firstSection);
            if (firstSectionRange != null)
                Marshal.ReleaseComObject(firstSectionRange);

            return firstSectionRangeNotDeleted;
        }


        /// <summary>
        ///     Returns a Word.ShapeRange from a Word.Range.
        ///     Unlike the Word supplied Word.ShapeRange member this method will not throw an exception if the ShapeRange does not
        ///     exist.
        /// </summary>
        /// <param name="source">A Range instance.</param>
        /// <returns>A Word.ShapeRange, or Null if the Word.Story does not exist</returns>

        // ReSharper disable once InconsistentNaming
        public static ShapeRange ShapeRangeLJ([NotNull] this Range source)
        {
            Check.NotNull(source, nameof(source));

            // ReSharper disable once SuspiciousTypeConversion.Global
            var hr = ((Range11)source).ShapeRange(out var shapeRange);
            return hr == 0 ? shapeRange : null;
        }


        /// <summary>
        ///     Returns a Word.WdStoryType from a Word.Range.
        ///     Unlike the Word supplied Word.StoryType member this method will not throw an exception.
        /// </summary>
        /// <param name="source">A WdStoryType value.</param>
        /// <returns>A WdStoryType value</returns>
        public static WdStoryType StoryTypeLJ([NotNull] this Range source)
        {
            Check.NotNull(source, nameof(source));
            // ReSharper disable once SuspiciousTypeConversion.Global
            var range11 = (Range11)source;

            // ReSharper disable once SuspiciousTypeConversion.Global
            var hr = range11.StoryType(out var st);
            return hr == 0 ? st : WdStoryType.wdCommentsStory;
        }

        /// <summary>
        /// Return true if the supplied range is in the MainRange or if the range is referenced in the main story.
        /// </summary>
        /// <param name="source">The source range.</param>
        /// /// <param name="rangeToTest">The range to test. If null false is returned.</param>
        /// <returns></returns>
        public static bool InMainStoryRange([NotNull] this Range source, Range rangeToTest)
        {
            Range reference;

            if (rangeToTest == null)
                return false;

            switch (source.StoryType)
            {
                case WdStoryType.wdMainTextStory:
                    return source.InRange(rangeToTest);
                case WdStoryType.wdEndnotesStory:
                    var endNotes = source.Endnotes;
                    var endNote = endNotes[1];
                    reference = endNote.Reference;
                    Marshal.ReleaseComObject(endNote);
                    Marshal.ReleaseComObject(endNotes);
                    break;
                case WdStoryType.wdFootnotesStory:
                    var footnotes = source.Footnotes;
                    var footnote = footnotes[1];
                    reference = footnote.Reference;
                    Marshal.ReleaseComObject(footnote);
                    Marshal.ReleaseComObject(footnotes);
                    break;
                case WdStoryType.wdCommentsStory:
                    var comments = source.Footnotes;
                    var comment = comments[1];
                    reference = comment.Reference;
                    Marshal.ReleaseComObject(comment);
                    Marshal.ReleaseComObject(comments);
                    break;
                default:
                    return source.InRange(rangeToTest);
            }

            var inRange = reference.InRange(rangeToTest);
            Marshal.ReleaseComObject(reference);
            return inRange;
        }

        public static Range MainStoryRange([NotNull] this Range source)
        {

            switch (source.StoryType)
            {
                case WdStoryType.wdMainTextStory:
                    return source;
                case WdStoryType.wdEndnotesStory:
                    var endnotes = source.Endnotes;
                    var endnote = endnotes[1];
                    var rng = endnote.Reference;
                    Marshal.ReleaseComObject(endnote);
                    Marshal.ReleaseComObject(endnotes);

                    return rng;
                case WdStoryType.wdFootnotesStory:
                    var footnotes = source.Footnotes;
                    var footnote = footnotes[1];
                    var rng2 = footnote.Reference;
                    Marshal.ReleaseComObject(footnote);
                    Marshal.ReleaseComObject(footnotes);
                    return rng2;

                case WdStoryType.wdCommentsStory:
                    var comments = source.Comments;
                    var comment = comments[1];
                    var rng3 = comment.Reference;
                    Marshal.ReleaseComObject(comment);
                    Marshal.ReleaseComObject(comments);
                    return rng3;
                default:
                    return source;
            }
        }


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


        /// <summary>
        ///     Returns a nullable Word.WdRectangleType from a Word.Rectangle.
        ///     This method will not thrown an exception if the property is not available.
        /// </summary>
        /// <param name="source">A Range instance.</param>
        /// <returns>A Word.WdRectangleType, or Nothing if the Word.WdRectangleType does not exist for the given Rectangle.</returns>
        // ReSharper disable once InconsistentNaming
        public static WdRectangleType? RectangleTypeLJ([NotNull] this Rectangle source)
        {
            Check.NotNull(source, nameof(source));

            // ReSharper disable once SuspiciousTypeConversion.Global
            var hr = ((Rectangle12)source).RectangleType(out var rt);
            return hr == 0 ? (WdRectangleType?)rt : null;
        }


        /// <summary>
        ///     Returns the first range that is not Protected, i.e. the user can edit.
        /// </summary>
        /// <param name="source">A Range instance.</param>
        /// <returns>null if the whole document is protected.</returns>

        // ReSharper disable once InconsistentNaming
        public static WordBoolean LJIsEditable([NotNull] this Range source)
        {
            Check.NotNull(source, nameof(source));

            var retVal = WordBoolean.Unknown;
            if (source.Document.ProtectionType == WdProtectionType.wdNoProtection)
            {
                retVal = WordBoolean.True;
            }
            else
            {
                // ReSharper disable once SuspiciousTypeConversion.Global
                var rng = (Range11)source;

                //4605 (PropertyNotAvailable) The Case method or property is not available because the document is locked for editing.
                //Note must use Case, using bold/italic etc. only works in 2003.
                var hr = rng.Case_Get(out _);
                switch (hr)
                {
                    case PropertyNotAvailable:
                        retVal = WordBoolean.False;
                        break;
                    case 0:
                        retVal = WordBoolean.True;
                        break;
                }
            }

            return retVal;
        }


        /// <summary>
        ///     Returns True if the input ranges overlap
        /// </summary>
        /// <param name="source">A Range instance.</param>
        /// <param name="range1">Word.Range object. Can be either collapsed or expanded. May not be Nothing.</param>
        /// <param name="range2">Word.Range object. Can be either collapsed or expanded. May not be Nothing.</param>
        /// <returns>Boolean. True if the ranges overlap.</returns>
        /// <remarks>
        ///     Special Cases:
        ///     If both ranges are expanded (Range.Start ? Range.End), then if they are contiguous, False is returned.
        ///     If one range is collapsed (Range.Start = Range.End) and it exists at the BEGINNING of the other range, True is
        ///     returned.
        ///     If one range is collapsed (Range.Start = Range.End) and it exists at the END of the other range, False is returned.
        /// </remarks>
        public static bool RangesOverlap([NotNull] this Application source, Range range1, Range range2)
        {
            Check.NotNull(source, nameof(source));
            Check.NotNull(range1, nameof(range1));
            Check.NotNull(range2, nameof(range2));

            if (range1.StoryType != range2.StoryType)
            {
                return false;
            }

            var start1 = range1.Start;
            var end1 = range1.End;
            var start2 = range2.Start;
            var end2 = range2.End;

            if ((start1 > end2) || (start2 > end1))
            {
                // They're dis-contiguous
                return false;
            }

            if (end1 == start2)
            {
                // Only deemed to overlap if range1 is an insertion point at the beginning of range2
                return end1 == start1;
            }

            if (end2 == start1)
            {
                // Only deemed to overlap if range2 is an insertion point at the beginning of range1
                return end2 == start2;
            }

            return true;
        }

        /// <summary>
        ///     Returns True if the input range overlaps with this Range
        /// </summary>
        /// <param name="source">A Range instance.</param>
        /// <param name="range">Word.Range object. Can be either collapsed or expanded. May not be Nothing.</param>
        /// <returns>Boolean. True if the ranges overlap.</returns>
        /// <remarks>
        ///     Special Cases:
        ///     If both ranges are expanded (Range.Start ? Range.End), then if they are contiguous, False is returned.
        ///     If one range is collapsed (Range.Start = Range.End) and it exists at the BEGINNING of the other range, True is
        ///     returned.
        ///     If one range is collapsed (Range.Start = Range.End) and it exists at the END of the other range, False is returned.
        /// </remarks>
        public static bool Overlaps([NotNull] this Range source, [NotNull] Range range)
        {
            Check.NotNull(source, nameof(source));
            Check.NotNull(range, nameof(range));

            return RangesOverlap(source.Application, range1: source, range2: range);
        }


        /// <summary>
        ///     Returns a Range within the Range provided.
        /// </summary>
        /// <param name="source">A Range instance.</param>
        /// <param name="startPos">
        ///     The start position of the sub range. if out side of the range provided, then this value will be
        ///     coerced to fix inside the range provided.
        /// </param>
        /// <param name="endPos">
        ///     the end position of the sub range. if out side of the range provided, then this value will be
        ///     coerced to fix inside the range provided.
        /// </param>
        /// <returns>A new range. constrained by the provided startPos and endPos</returns>
        /// <remarks>Useful for debugging.</remarks>
        public static Range SubRange([NotNull] this Range source, int startPos, int endPos)
        {
            Check.NotNull(source, nameof(source));

            if (startPos < source.Start)
                startPos = source.Start;
            if (endPos > source.End)
                endPos = source.End;

            var rng = source.Duplicate;
            rng.SetRange(startPos, endPos);


            return rng;
        }


        /// <summary>
        ///     Returns information about the specified selection or range.
        /// </summary>
        /// <param name="source">A Range instance.</param>
        /// <param name="type">Required WdInformation. The information type.</param>
        /// <returns>The value returned is dependent of the type of information requested.</returns>
        public static object Information([NotNull] this Range source, WdInformation type)
        {
            Check.NotNull(source, nameof(source));
            // ReSharper disable once UseIndexedProperty
            return source.get_Information(type);
        }


        public static int Length([NotNull] this Range source) => source.End - source.Start;


        //Word.Range.12 Content Controls


        /// <summary>
        ///     Returns a list of Word.Range objects that encapsulate each ContentControl object within a Word.Range
        /// </summary>
        /// <param name="source">A Range instance.</param>
        /// <returns>A list of Word.Range objects, or null if the version of Word is less than version 12.</returns>
        /// <remarks>Only valid in Office 12 and beyond. Will not throw an exception if used on an earlier version of Word.</remarks>
        // ReSharper disable once InconsistentNaming
        public static IEnumerable<Range> ContentControlRangesLJ([NotNull] this Range source)
        {
            Check.NotNull(source, nameof(source));

            foreach (ContentControl cc in source.ContentControls)
            {
                var range = cc.Range;
                Marshal.ReleaseComObject(cc);
                yield return range;
            }
        }


        /// <summary>
        ///     Returns a list of Word.Revision objects that truly overlap a Word.Range.
        ///     This is done because Word returns a collection where there are more elements
        ///     than the count (e.g. .Count = 0, but For/Next returns multiple entries)
        /// </summary>
        /// <param name="source">A Word.Range instance.</param>
        /// <remarks>
        ///     IMPORTANT: Because of a bug in Word 2013 and later, Word.Range.Revisions
        ///     should never be used. For any particular range object, accessing the Revisions
        ///     corrupts the collection. This extension provides a workaround for this problem.
        /// </remarks>
        /// <returns>A list of Word.Revision objects.</returns>
        // ReSharper disable once InconsistentNaming
        public static IEnumerable<Revision> RevisionsLJ([NotNull] this Range source)
        {
            Check.NotNull(source, nameof(source));

            //	KDP Comments
            // 01 Oct 2015: Word 2013 has a defect whereby if you set a pointer to the Revisions
            //							 object, you will get the correct collection only the first time
            //							 you read the collection. After that, the collection is wrong.
            //							 EXAMPLE: Open document from TSWA 5275 and go into the VB Editor,
            //							 Immediate window. Issue the following commands (return vals in [])
            //
            //	>set o = ActiveDocument.Footnotes(9).Reference.Revisions [(no return value)]
            // >?o.Count [1]
            // >?o.Count [414]
            //	
            //							The workaround is to cast the Revisions to a static List, then
            //							use the list for getting counts and for the loop enumeration.
            //							NOTE: The problem is not fixed in 2016 RTM.

            // Note: Count is number of true revs, so if count = 0 there are none;
            //       however, ther can be additional members returned in collection
            //       before it gets to the true ones.
            foreach (Revision rev in source.Revisions)
            {
                var rng = rev.Range;


                if (rng.End <= source.Start || rng.Start >= source.End)
                {
                    Marshal.ReleaseComObject(rng);
                    Marshal.ReleaseComObject(rev);
                    continue;
                }

                yield return rev;
            }
        }


        public static string EmphasisDisplayText(this Range source, CharacterFormattingEmphasis value)
        {
            var result = string.Empty;

            var emphasis = value & CharacterFormattingEmphasis.Any;
            if (emphasis == CharacterFormattingEmphasis.None)
            {
                result = "None";
            }
            else
            {
                if ((emphasis & CharacterFormattingEmphasis.Underline) == CharacterFormattingEmphasis.Underline)
                {
                    result += "U";
                }

                if ((emphasis & CharacterFormattingEmphasis.Italic) == CharacterFormattingEmphasis.Italic)
                {
                    result += "I";
                }

                if ((emphasis & CharacterFormattingEmphasis.SmallCaps) == CharacterFormattingEmphasis.SmallCaps)
                {
                    result += "S";
                }
            }

            return result;
        }


        /// <summary>Applies specialized emphasis to a Word source to set the character formatting in that source.</summary>
        /// <param name="source">Range to be applied with the CharacterFormattingEmphasis.</param>
        /// <param name="emphasisString">Text containing emphasis bits to apply to the source.</param>
        public static void ApplyEmphasisString([NotNull] this Range source, string emphasisString)
        {
            CheckNotNull(source);

            var limit = Convert.ToInt32(Math.Min(emphasisString.Length, source.End - source.Start)) - 1;
            if (limit < 0)
            {
                // One of the parameters was zero length
                return;
            }

            var rng = source.Duplicate;
            rng.Collapse();
            var nextPos = 0;

            do // Process one span
            {
                // See how many characters have identical emphasis
                var emphasisChar = emphasisString[nextPos];
                var len = 1;
                nextPos += 1;
                while (!(nextPos > limit || emphasisString[nextPos] != emphasisChar))
                {
                    len += 1;
                    nextPos += 1;
                }

                // Create a source of that length
                rng.MoveEnd(WdUnits.wdCharacter, len);

                // And apply that emphasis
                if (int.TryParse(emphasisChar.ToString(), out var intEmph))
                {
                    var emph = (CharacterFormattingEmphasis)intEmph;
                    rng.Font.SmallCaps = Convert.ToInt32((emph & CharacterFormattingEmphasis.SmallCaps) == CharacterFormattingEmphasis.SmallCaps);
                    rng.Font.Italic = Convert.ToInt32((emph & CharacterFormattingEmphasis.Italic) == CharacterFormattingEmphasis.Italic);
                    rng.Font.Underline = Convert.ToBoolean((emph & CharacterFormattingEmphasis.Underline) == CharacterFormattingEmphasis.Underline) ? WdUnderline.wdUnderlineSingle : WdUnderline.wdUnderlineNone;
                }

                rng.Collapse(WdCollapseDirection.wdCollapseEnd);
            } while (!(nextPos > limit));

            Marshal.ReleaseComObject(rng);
        }


        /// <summary>
        ///     Returns the Style of the supplied source Range.
        /// </summary>
        /// <param name="source">A Range instance.</param>
        /// <returns>A Style instance.</returns>
        public static Style Style([NotNull] this Range source)
        {
            Check.NotNull(source, nameof(source));

            // ReSharper disable once UseIndexedProperty
            return (Style)source.get_Style();
        }

        /// <summary>
        ///     Sets the Style of the supplied source Range.
        /// </summary>
        /// <param name="source">A Range instance.</param>
        /// <param name="styleName">The name of the style to apply.</param>
        public static void Style([NotNull] this Range source, [NotNull] string styleName)
        {
            Check.NotNull(source, nameof(source));
            Check.NotNull(styleName, nameof(styleName));

            // ReSharper disable once UseIndexedProperty
            source.set_Style(styleName);
        }

        /// <summary>
        ///     Sets the Style of the supplied source Range.
        /// </summary>
        /// <param name="source">A Range instance.</param>
        /// <param name="style">The Style instance to apply.</param>
        public static void Style([NotNull] this Range source, [NotNull] Style style)
        {
            Check.NotNull(source, nameof(source));
            Check.NotNull(style, nameof(style));

            // ReSharper disable once UseIndexedProperty
            source.set_Style(style);
        }


        /// <summary>
        ///     Calls Window.SeekView for applicable to the story type for the range and then selects the range.
        /// </summary>
        /// <param name="source">A Range instance.</param>
        /// <param name="collapseDirection"></param>
        public static void SelectAndSeek(this Range source, WdCollapseDirection? collapseDirection = null)
        {
            if (collapseDirection != null)
                source.Collapse(collapseDirection.GetValueOrDefault());

            TrySeekView(source);
            source.Select();
        }


        /// <summary>
        ///     Calls Window.SeekView for applicable to the story type for the range.
        /// </summary>
        /// <param name="source">A Range instance.</param>
        public static void TrySeekView(this Range source)
        {
            var view = source.Document.ActiveWindow.View;
            if (view.Type != WdViewType.wdPrintView)
                return;

            var newSeekView = view.SeekView;
            switch (source.StoryType)
            {
                case WdStoryType.wdMainTextStory:
                    newSeekView = WdSeekView.wdSeekMainDocument;
                    break;
                case WdStoryType.wdFootnotesStory:
                    newSeekView = WdSeekView.wdSeekFootnotes;
                    break;
                case WdStoryType.wdEndnotesStory:
                    newSeekView = WdSeekView.wdSeekEndnotes;
                    break;
            }

            if (view.SeekView != newSeekView)
                view.SeekView = newSeekView;

            Marshal.ReleaseComObject(view);

        }



        /// <summary></summary>
        /// <param name="source"></param>
        /// <param name="textToFind"></param>
        /// <param name="ignoreCase"></param>
        /// <param name="direction"></param>
        /// <param name="useNormalizedText"></param>

        public static Range FindSubRange([NotNull] this Range source, [NotNull] string textToFind, bool ignoreCase, WdConstants direction,
                                         bool useNormalizedText)
        {
            CheckNotNull(source);
            Check.NotNull(textToFind, nameof(textToFind));
            if (string.IsNullOrEmpty(source.Text))
            {
                return null;
            }

            if (string.IsNullOrEmpty(textToFind))
            {
                return null;
            }

            var oDupRange = source.Duplicate;
            oDupRange.TextRetrievalMode.IncludeFieldCodes = source.TextRetrievalMode.IncludeFieldCodes;
            oDupRange.TextRetrievalMode.IncludeHiddenText = source.TextRetrievalMode.IncludeHiddenText;

            var startWork = oDupRange.Start;
            var endWork = oDupRange.End;

            var workRangeText = oDupRange.Text;
            var workTextToFind = textToFind;
            if (useNormalizedText)
            {
                workRangeText = workRangeText.Clean();
                workTextToFind = workTextToFind.Clean();
            }

            var comparison = ignoreCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            var lInstr = direction == WdConstants.wdForward
                             ? workRangeText.IndexOf(workTextToFind, comparison)
                             : workRangeText.LastIndexOf(workTextToFind, comparison);

            if (workRangeText == workTextToFind)
            {
                return oDupRange;
            }

            if (lInstr == -1)
            {
                return null;
            }

            if (textToFind.Length > 255)
            {
                // Look for string directly
                oDupRange.Start = oDupRange.Start + lInstr;
                oDupRange.End = oDupRange.Start + workTextToFind.Length;
                return oDupRange;
            }

            if (source.Fields.Count == 0)
            {
                // Use direct computation of start and end points of subrange
                oDupRange.Start = startWork + lInstr;
                oDupRange.End = oDupRange.Start + workTextToFind.Length;
                return oDupRange;
            }

            var lLoopEnd = direction == WdConstants.wdForward ? 1 : workRangeText.CountOf(workTextToFind);

            // Doing search, don't use normalized text
            var origBrowserTarget = source.Application.Browser.Target;
            oDupRange.Collapse(WdCollapseDirection.wdCollapseStart);
            for (var idx = 1; idx <= lLoopEnd; idx++)
            {
                oDupRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                var findText = (object)textToFind;
                var find = oDupRange.Find;
                find.Execute(findText, MatchCase: !ignoreCase, Forward: true, Wrap: WdFindWrap.wdFindStop);
                if (idx != lLoopEnd) continue;
                if (string.Equals(oDupRange.Text, textToFind, StringComparison.CurrentCultureIgnoreCase) &&
                    oDupRange.Start >= startWork && oDupRange.End <= endWork)
                {
                    return oDupRange;
                }

                Marshal.ReleaseComObject(find);
            }

            if (origBrowserTarget >= WdBrowseTarget.wdBrowsePage &&
                source.Application.Browser.Target != origBrowserTarget) // wdBrowsePage = 1
            {
                source.Application.Browser.Target = origBrowserTarget;
            }

            return null;
        }


        /// <summary>Finds first contiguous source of underlined text within a source.</summary>
        /// <param name="source">Range to search for underlined text.</param>
        /// <returns>Range containing underlined text.</returns>
        public static Range FindFirstUnderlinedSubrange([NotNull] this Range source)
        {
            CheckNotNull(source);
            if (source.Start == source.End)
                return null;

            var sourceFont = source.Font;
            if (sourceFont.Underline == WdUnderline.wdUnderlineNone)
            {
                Marshal.ReleaseComObject(sourceFont);
                return null;
            }

            var dupRange = source.Duplicate;
            dupRange.Collapse(WdCollapseDirection.wdCollapseStart);
            dupRange.End += 1;
            var dupFont = dupRange.Font;
            while (dupFont.Underline == WdUnderline.wdUnderlineNone && dupRange.End < source.End)
            {
                Marshal.ReleaseComObject(dupFont);
                dupRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                dupRange.End += 1;
                dupFont = dupRange.Font;
            }

            Marshal.ReleaseComObject(dupFont);
            dupFont = dupRange.Font;

            while (dupFont.Underline == WdUnderline.wdUnderlineSingle && dupRange.End < source.End)
            {
                Marshal.ReleaseComObject(dupFont);
                dupRange.End += 1;
                dupFont = dupRange.Font;
            }

            Marshal.ReleaseComObject(dupFont);
            dupFont = dupRange.Font;

            if (dupFont.Underline != WdUnderline.wdUnderlineSingle)
            {
                Marshal.ReleaseComObject(dupFont);
                dupRange.End -= 1;
                dupFont = dupRange.Font;
            }
            Marshal.ReleaseComObject(dupFont);
            if (dupRange.Start != source.End)
                return dupRange;

            Marshal.ReleaseComObject(dupRange);
            return null;

        }


        /// <summary>
        ///     Trims the characters provided from the start of the Range. If charactersToTrim is null all which space characters
        ///     are trimmed.
        /// </summary>
        /// <param name="source">The source Range to trim</param>
        /// <param name="charachtersToTrim">
        ///     A string containing the characters to trim. If null the all white space characters are
        ///     removed.
        /// </param>
        /// <param name="removeCharacters">True to remove the characters from the range false to only adjust the source range.</param>
        /// <param name="trimOutsideofRange">
        ///     True to allow the source range to continue to remove characters beyond is current
        ///     start and end positions.
        /// </param>
        public static void TrimStart([NotNull] this Range source, string charachtersToTrim = null, bool removeCharacters = false, bool trimOutsideofRange = false)
        {
            TrimCore(source, charachtersToTrim, removeCharacters, trimOutsideofRange, trimStart: true, trimEnd: false);
        }


        /// <summary>
        ///     Trims the characters provided from the end of the Range. If charactersToTrim is null all which space characters are
        ///     trimmed.
        /// </summary>
        /// <param name="source">The source Range to trim</param>
        /// <param name="charactersToTrim">
        ///     A string containing the characters to trim. If null the all white space characters are
        ///     removed.
        /// </param>
        /// <param name="removeCharacters">True to remove the characters from the range false to only adjust the source range.</param>
        /// <param name="trimOutsideOfRange">
        ///     True to allow the source range to continue to remove characters beyond is current
        ///     start and end positions.
        /// </param>
        public static void TrimEnd([NotNull] this Range source, string charactersToTrim = null, bool removeCharacters = false, bool trimOutsideOfRange = false)
        {
            TrimCore(source, charactersToTrim, removeCharacters, trimOutsideOfRange, trimStart: false, trimEnd: true);
        }


        /// <summary>
        ///     Trims the characters provided from the Range. If charactersToTrim is null all which space characters are trimmed.
        /// </summary>
        /// <param name="source">The source Range to trim, on return the source range bounds are adjusted to the trimmed size.</param>
        /// <param name="charachtersToTrim">
        ///     A string containing the characters to trim. If null the all white space characters are
        ///     removed.
        /// </param>
        /// <param name="removeCharacters">True to remove the characters from the range false to only adjust the source range.</param>
        /// <param name="trimOutsideofRange">
        ///     True to allow the source range to continue to remove characters beyond is current
        ///     start and end positions.
        /// </param>
        public static void Trim([NotNull] this Range source, string charachtersToTrim = null, bool removeCharacters = false, bool trimOutsideofRange = false)
        {
            TrimCore(source, charachtersToTrim, removeCharacters, trimOutsideofRange, trimStart: true, trimEnd: true);
        }


        private static void TrimCore(Range source, string charactersToTrim, bool removeCharacters, bool trimOutsideOfRange, bool trimStart, bool trimEnd)
        {
            CheckNotNull(source);

            if (source.Start == source.End)
                return;

            var moved = false;
            int newPos;
            var refRange = removeCharacters ? source.Duplicate : null;
            var start = source.Start;
            var end = source.End;

            if (trimEnd)
            {
                if (trimOutsideOfRange && removeCharacters)
                {
                    // Check after end of range
                    moved = source.MoveEndWhile(charactersToTrim, WdConstants.wdForward) != 0;
                    newPos = source.End; // Temp store while we update
                    source.End = end; // Set to old end so we don't compare more than we need.
                    end = newPos; // Set new end bounds
                }

                // Note wdBackward returns negative counts
                moved |= source.MoveEndWhile(charactersToTrim, WdConstants.wdBackward) != 0;

                if (trimOutsideOfRange == false && source.End < start)
                {
                    source.SetRange(start, start); // Fix the range bounds if MoveStartWhile moved outside the range
                    trimStart = false; // Can exit early in this case.
                }

                if (removeCharacters && moved)
                {
                    // Delete start whitespace characters. Don't use rng.Delete is its buggy and Deletes extra characters
                    refRange.SetRange(source.End, end);
                    if (refRange.Text != null)
                        refRange.Text = null;
                }
            }

            if (!trimStart)
                return;

            moved = false;

            if (trimOutsideOfRange && removeCharacters)
            {
                // Check before start of range
                // Note wdBackward returns negative counts
                moved = source.MoveStartWhile(charactersToTrim, WdConstants.wdBackward) != 0;
                newPos = source.Start; // Temp store while we update
                source.Start = start; // Set to old start so we don't compare more than we need.
                start = newPos; // Set new end bounds
            }

            moved |= source.MoveStartWhile(charactersToTrim, WdConstants.wdForward) != 0;
            if (trimOutsideOfRange == false && source.Start > end)
                source.SetRange(end, end); //Fix the range bounds if MoveStartWhile moved outside the range

            if (removeCharacters && moved)
            {
                // Delete end whitespace characters. Don't use rng.Delete is its buggy and Deletes extra characters
                refRange.SetRange(start, source.Start);
                if (refRange.Text != null)
                    refRange.Text = null;
            }
        }


        public static void TrimSectionBreak([NotNull] this Range source)
        {
            var workRange = source.Duplicate;
            workRange.Collapse(WdCollapseDirection.wdCollapseStart);
            workRange.End += 1;
            if (workRange.IsSectionBreak())
                source.Start += 1;
            workRange.SetRange(source.End, source.End);
            workRange.Start -= 1;
            if (workRange.IsSectionBreak())
                source.End -= 1;
            Marshal.ReleaseComObject(workRange);
        }


        [CLSCompliant(false)]
        // ReSharper disable once InconsistentNaming
        public static Page PageLJ([NotNull] this Range source)
        {
            Check.NotNull(source, nameof(source));

            var startPos = source.Start;
            var sourceStoryType = source.StoryType;
            foreach (Page pg in source.Document.ActiveWindow.Panes[1].Pages)
            {
                if ((pg.Rectangles.Cast<Rectangle>().Select(rect => rect.RangeLJ())).Any(rectRange => rectRange != null
                                                                                                      && rectRange.StoryType == sourceStoryType
                                                                                                      && rectRange.Start <= startPos
                                                                                                      && rectRange.End >= startPos))
                {
                    return pg;
                }
            }

            return null;
        }


        /// <summary>Determines if a given Word source is fully contained within a field.</summary>
        /// <returns>Returns true if the source is fully contained within a field.</returns>
        public static bool InField([NotNull] this Range source)
        {
            CheckNotNull(source);
            if (source.Application.VersionLJ() >= OfficeVersion.Office2013)
            {
                //Use new WdInformation flags 
                const int wdInFieldCode = 44;
                const int wdInFieldResult = 45;

                if (!((bool)source.Information((WdInformation)wdInFieldCode) && (bool)source.Information((WdInformation)wdInFieldResult)))
                    return false; // Not in any fields
            }

            return FieldContainersInternal(source).Any();
        }


        /// <summary>Determines field that fully contains a given Word source.</summary>
        /// <returns>Returns a field object that fully contains the input Word source. Returns null if there isn't one.</returns>
        public static IEnumerable<Field> FieldContainers([NotNull] this Range source)
        {
            CheckNotNull(source);
            if (source.Application.VersionLJ() >= OfficeVersion.Office2013)
            {
                //Use new WdInformation flags 
                const int wdInFieldCode = 44;
                const int wdInFieldResult = 45;

                if (!((bool)source.Information((WdInformation)wdInFieldCode) && (bool)source.Information((WdInformation)wdInFieldResult)))
                    return Enumerable.Empty<Field>(); // Not in any fields
            }

            return FieldContainersInternal(source);
        }

        private static IEnumerable<Field> FieldContainersInternal(Range source)
        {
            var hasFields = false;
            foreach (Field fld in source.Fields)
            {
                var fldResult = fld.ResultLJ();
                var endPos = fldResult?.End ?? fld.Code.End;
                if (source.Start >= fld.Code.Start && source.End <= endPos)
                    yield return fld;
            }
        }

        /// <summary>Determines if a given Word source is fully contained within a table.</summary>
        /// <returns>Returns true if the source is fully contained within a table.</returns>
        public static bool InTable([NotNull] this Range source)
        {
            CheckNotNull(source);
            return (bool)source.Information(WdInformation.wdWithInTable);
        }


        /// <summary>Determines table that fully contains a given Word source.</summary>
        /// <param name="source">Input Word source.</param>
        /// <returns>Returns a table object that fully contains the input Word source. Returns null if there isn't one.</returns>
        public static Table TableContainer([NotNull] this Range source)
        {
            CheckNotNull(source);

            if ((bool)source.Information(WdInformation.wdWithInTable))
            {
                return source.Tables[1];
            }

            return null;
            //OLDCODE
            //TODO cut-over check it is the same
            //         var refStory = source.StoryType;
            //var refStart = source.Start;
            //var refEnd = source.End;
            //var rangeIsCollapsed = refStart == refEnd;

            //foreach (Table tbl in source.Document.Tables)
            //{
            //	if (tbl.Range.StoryType != refStory)
            //                 continue;

            //	if (rangeIsCollapsed)
            //	{
            //		if (refStart >= tbl.Range.Start && refEnd < tbl.Range.End)
            //			return tbl;
            //	}
            //	else if (refStart >= tbl.Range.Start && refEnd <= tbl.Range.End)
            //	{
            //		return tbl;
            //	}
            //}

            //return null;
        }


        public static bool InTableOfContents([NotNull] this Range source)
        {
            CheckNotNull(source);
            return source.TableOfContentsContainer() != null;
        }


        /// <summary>Gets the Word.TableOfContents to which a source belongs.</summary>
        /// <returns>The containing TableOfContents.</returns>
        public static TableOfContents TableOfContentsContainer([NotNull] this Range source)
        {
            CheckNotNull(source);
            //NJKA: Changed to range logic, from Overlaps. If we are at the end of the range Overlaps returns false, however we are in the TOC.
            //This saves altering the behavior of the Range.Overlaps method
            //return source.Document.TablesOfContents.Cast<TableOfContents>().FirstOrDefault(toc => toc.Range.Overlaps(source) || toc.Range.End == source.End);
            return source.Document.TablesOfContents.Cast<TableOfContents>().FirstOrDefault(toc => source.Start >= toc.Range.Start &&
                                                                                                  source.End <= toc.Range.End);
        }


        public static bool InTableOfAuthorities([NotNull] this Range source)
        {
            CheckNotNull(source);
            return source.TableOfAuthoritiesContainer() != null;
        }


        /// <summary>Gets the Word.TableOfAuthorities to which a source belongs.</summary>
        /// <returns>The containing TableOfAuthorities.</returns>
        public static TableOfAuthorities TableOfAuthoritiesContainer([NotNull] this Range source)
        {
            CheckNotNull(source);
            //NJKA: Changed to range logic, from Overlaps. If we are at the end of the range Overlaps returns false, however we are in the TOA.
            //This saves altering the behavior of the Range.Overlaps method
            //return source.Document.TablesOfContents.Cast<TableOfAuthorities>().FirstOrDefault(toc => toc.Range.Overlaps(source) || toc.Range.End == source.End);
            return source.Document.TablesOfAuthorities.Cast<TableOfAuthorities>().FirstOrDefault(toc => source.Start >= toc.Range.Start &&
                                                                                                        source.End <= toc.Range.End);
        }


        public static bool InTableOfFigures([NotNull] this Range source)
        {
            CheckNotNull(source);
            return source.TableOfFiguresContainer() != null;
        }


        /// <summary>Gets the Word.TableOfFigures to which a source belongs.</summary>
        /// <returns>The containing TableOfFigures.</returns>
        public static TableOfFigures TableOfFiguresContainer([NotNull] this Range source)
        {
            CheckNotNull(source);

            return source.Document.TablesOfFigures.Cast<TableOfFigures>().FirstOrDefault(toc => toc.Range.Overlaps(source));
        }


        /// <summary>Returns Page rectangle containing the input range, if there is one.</summary>
        [CLSCompliant(false)]
        public static object ContainingPageRectangle(this Range source)
        {
            var wordDoc = source.Document;
            var pn = wordDoc.ActiveWindow.ActivePane;
            var pg = pn.Pages[Convert.ToInt32(source.Information(WdInformation.wdActiveEndPageNumber))];
            foreach (Rectangle rect in pg.Rectangles)
            {
                var rectType = rect.RectangleTypeLJ();
                if (rectType == null || rectType != WdRectangleType.wdDocumentControlRectangle)
                    continue;

                var rectRange = rect.RangeLJ();
                if (rectRange == null)
                    continue;

                if (rectRange.Overlaps(source))
                {
                    return rect;
                }
            }

            return null;
        }


        /// <summary>Determines if a Word source is a source break.</summary>
        /// <returns>Returns true if break in the source is a source break.</returns>
        public static bool IsSectionBreak([NotNull] this Range source)
        {
            CheckNotNull(source);

            var rangeText = source.Text;

            // Range must contain one and only one character
            // and that character must be an ASCII 12 character.
            if (string.IsNullOrEmpty(rangeText))
            {
                return false;
            }

            if (rangeText.Length != 1)
            {
                return false;
            }

            if (rangeText != Convert.ToChar(value: 12).ToString())
            {
                return false;
            }

            // A character sourceText of 12 in Word can be either a page break or a source break
            // Break is a source break if ranges on opposite sides
            // of the break are in different sections.

            var rngWork = source.Duplicate;
            rngWork.Collapse(WdCollapseDirection.wdCollapseStart);
            var sectionIndex1 = rngWork.Sections[Index: 1].Index;

            rngWork = source.Duplicate;

            rngWork.Collapse(WdCollapseDirection.wdCollapseEnd);
            var sectionIndex2 = rngWork.Sections[Index: 1].Index;

            return sectionIndex1 != sectionIndex2;
        }


        /// <summary>Determines if a range preceeds specified character, optionally ignoring casing and/or white space.</summary>
        /// <returns>Returns true if range preceeds input character.</returns>
        public static bool Precedes([NotNull] this Range source, char c, bool caseSensitive, bool ignoreWhiteSpace)
        {
            return source.Precedes(c.ToString(), caseSensitive, ignoreWhiteSpace);
        }

        /// <summary>Determines if a range precedes specified text, optionally ignoring casing and/or white space.</summary>
        /// <returns>Returns true if range precedes input text.</returns>
        public static bool Precedes([NotNull] this Range source, string text, bool caseSensitive, bool ignoreWhiteSpace)
        {
            CheckNotNull(source);
            Check.NotNull(text, nameof(text));
            //Check.NotEmpty(text, nameof(text));

            var workRange = source.Duplicate;
            workRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            if (ignoreWhiteSpace)
            {
                char curChar;
                int rangeEnd;
                do
                {
                    workRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    rangeEnd = workRange.End;
                    if (workRange.MoveEnd(WdUnits.wdCharacter, 1) == 0)
                        break;

                    curChar = Convert.ToChar(workRange.Characters[1].Text);
                } while (char.IsWhiteSpace(curChar) && workRange.End > rangeEnd);

                workRange.Collapse(WdCollapseDirection.wdCollapseStart);
            }

            var numCharsMoved = workRange.MoveEnd(WdUnits.wdCharacter, text.Length);
            if (numCharsMoved != text.Length) return false;
            var workRangeText = caseSensitive ? workRange.Text : workRange.Text.ToLower();
            text = caseSensitive ? text : text.ToLower();
            return workRangeText == text;
        }


        /// <summary>Determines if a range follows specified character, optionally ignoring casing and/or white space.</summary>
        /// <returns>Returns true if range follows input character.</returns>
        public static bool Follows([NotNull] this Range source, char c, bool caseSensitive, bool ignoreWhiteSpace)
        {
            return source.Follows(c.ToString(), caseSensitive, ignoreWhiteSpace);
        }

        /// <summary>Determines if a range follows specified text, optionally ignoring casing and/or white space.</summary>
        /// <returns>Returns true if range follows input text.</returns>
        public static bool Follows([NotNull] this Range source, string text, bool caseSensitive, bool ignoreWhiteSpace)
        {
            CheckNotNull(source);
            Check.NotNull(text, nameof(text));

            var workRange = source.Duplicate;
            workRange.Collapse(WdCollapseDirection.wdCollapseStart);
            if (ignoreWhiteSpace)
            {
                int rangeStart;
                char curChar;
                do
                {
                    workRange.Collapse(WdCollapseDirection.wdCollapseStart);
                    rangeStart = workRange.Start;
                    if (workRange.MoveStart(WdUnits.wdCharacter, -1) == 0)
                        break;

                    curChar = Convert.ToChar(workRange.Characters[1].Text);
                } while (char.IsWhiteSpace(curChar) && workRange.Start < rangeStart);

                workRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            }

            var numCharsMoved = workRange.MoveStart(WdUnits.wdCharacter, -1 * text.Length);
            if (numCharsMoved != -1 * text.Length) return false;
            var workRangeText = caseSensitive ? workRange.Text : workRange.Text.ToLower();
            text = caseSensitive ? text : text.ToLower();
            return workRangeText == text;
        }


        /// <summary>Does a format-sensitive text replacement of text within a Word source.</summary>
        /// <param name="source">Reference Word source.</param>
        /// <param name="textToReplace">Search text.</param>
        /// <param name="replacementText">Replacement text.</param>
        /// <param name="matchCase">If true, do not replace text whose case does not match textToReplace.</param>
        /// <param name="matchWholeWord">If true, match on whole words only.</param>
        /// <param name="validateRange">If true, does a second check to ensure that the found text matches textToReplace.</param>
        /// <param name="maxCountForValidate">Number of mismatches of found text allowed before exiting.</param>
        /// <param name="resetReplacementFont">If true, resets font formatting of replaced text.</param>
        /// <returns>Returns true if search text was found and replacement successful.</returns>
        public static bool Replace([NotNull] this Range source, [NotNull] string textToReplace, [CanBeNull] string replacementText,
                                   bool matchCase, bool matchWholeWord, bool validateRange,
                                   long maxCountForValidate, bool resetReplacementFont)
        {
            // Parameter validations
            Check.NotNull(source, nameof(source));
            Check.NotNull(textToReplace, nameof(textToReplace));

            if (string.IsNullOrEmpty(source.Text))
            {
                return false;
            }

            if (textToReplace.Length > 255)
            {
                throw new ArgumentOutOfRangeException(nameof(textToReplace));
            }

            if (replacementText != null && replacementText.Length > 255)
            {
                throw new ArgumentOutOfRangeException(nameof(replacementText));
            }

            if (replacementText == null)
            {
                replacementText = string.Empty;
            }

            var doc = source.Document;

            var textToFindIdx = source.Text.IndexOf(textToReplace, startIndex: 0,
                                                    comparisonType:
                                                    matchCase
                                                        ? StringComparison.InvariantCulture
                                                        : StringComparison.InvariantCultureIgnoreCase);
            if (textToFindIdx < 0 && !textToReplace.StartsWith("^"))
            {
                return false;
            }

            // Find does not work well with vbCr as the ReplacementText
            // Use special control characters instead
            if (textToReplace == Environment.NewLine)
            {
                textToReplace = "^p";
            }
            else if (textToReplace == Convert.ToChar(value: 10).ToString() ||
                     textToReplace == Convert.ToChar(value: 11).ToString())
            {
                textToReplace = "^l";
            }

            var beforeText = source.Text;

            var origBrowserTarget = source.Application.Browser.Target;
            if (validateRange)
            {
                var foundCount = 0;
                var loopDone = false;
                var workRange = source.Duplicate;
                workRange.Collapse(WdCollapseDirection.wdCollapseStart);
                while (!loopDone)
                {
                    workRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    workRange.Find.ClearFormatting();
                    workRange.Find.Text = textToReplace;
                    workRange.Find.Format = false;
                    workRange.Find.MatchCase = matchCase;
                    workRange.Find.MatchWholeWord = matchWholeWord;
                    workRange.Find.Wrap = WdFindWrap.wdFindStop;
                    workRange.Find.Execute();
                    if (workRange.Find.Found & workRange.InRange(source))
                    {
                        workRange.Text = replacementText;
                        if (resetReplacementFont)
                        {
                            workRange.Font.Reset();
                        }
                    }
                    else
                    {
                        loopDone = true;
                    }

                    foundCount = foundCount + 1;
                    loopDone = loopDone || (foundCount >= maxCountForValidate);
                }
            }
            else
            {
                if (replacementText == Convert.ToChar(value: 13).ToString())
                {
                    replacementText = "^p";
                }
                else if (replacementText == "\v")
                {
                    replacementText = "^l";
                }

                source.Find.ClearFormatting();
                source.Find.Replacement.ClearFormatting();
                source.Find.Text = textToReplace;
                source.Find.Format = false;
                source.Find.Forward = true;
                source.Find.Replacement.Text = replacementText;
                if (resetReplacementFont)
                {
                    source.Find.Replacement.Font.Bold = 0;
                    source.Find.Replacement.Font.Italic = 0;
                    source.Find.Replacement.Font.Underline = WdUnderline.wdUnderlineNone;
                }

                source.Find.MatchCase = matchCase;
                source.Find.MatchWholeWord = matchWholeWord;
                source.Find.MatchWildcards = false;
                source.Find.MatchSoundsLike = false;
                source.Find.MatchAllWordForms = false;
                source.Find.Wrap = WdFindWrap.wdFindStop;
                source.Find.Execute(Replace: WdReplace.wdReplaceAll);
            }

            if (origBrowserTarget >= WdBrowseTarget.wdBrowsePage &&
                source.Application.Browser.Target != origBrowserTarget) // wdBrowsePage = 1
            {
                source.Application.Browser.Target = origBrowserTarget;
            }


            var afterText = source.Text;

            if (afterText == beforeText)
            {
                //Debug.Assert(false, "WordHelper.ReplaceTextWithinRange: Could not perform automatic text replacement");

                Range dupRange = null;

                var loopDone = false;
                while (!loopDone)
                {
                    textToFindIdx = source.Text.IndexOf(textToReplace, startIndex: 0,
                                                        comparisonType:
                                                        matchCase
                                                            ? StringComparison.InvariantCulture
                                                            : StringComparison.InvariantCultureIgnoreCase);
                    if (textToFindIdx > 0)
                    {
                        var rngStart = source.Start + textToFindIdx - 1;
                        switch (source.StoryType)
                        {
                            case WdStoryType.wdMainTextStory:
                                var start = (object)rngStart;
                                dupRange = doc.Range(start, start);
                                break;
                            case WdStoryType.wdFootnotesStory:
                            case WdStoryType.wdEndnotesStory:
                                dupRange = source.Duplicate;
                                dupRange.Collapse(WdCollapseDirection.wdCollapseStart);
                                dupRange.Start = rngStart;
                                break;
                        }

                        dupRange.End = rngStart + textToReplace.Length;
                        bool doTheReplace;
                        if (matchCase)
                        {
                            doTheReplace = Convert.ToBoolean(dupRange.Text == textToReplace);
                        }
                        else
                        {
                            doTheReplace =
                                string.Compare(dupRange.Text, textToReplace, StringComparison.OrdinalIgnoreCase) == 0;
                        }

                        var beforeLoopText = source.Text;
                        if (doTheReplace)
                        {
                            try
                            {
                                dupRange.Text = replacementText;
                            }
                            catch (Exception)
                            {
                                // Ignore the error
                            }

                            if (resetReplacementFont)
                            {
                                if (dupRange.Text == replacementText)
                                {
                                    dupRange.Font.Reset();
                                }
                            }
                        }

                        loopDone = Convert.ToBoolean(source.Text == beforeLoopText);
                    }
                    else
                    {
                        loopDone = true;
                    }
                }
            }

            afterText = source.Text;

            return Convert.ToBoolean(afterText != beforeText);
        }


        /// <summary>Does a plain-text replacement of text within a Word source.</summary>
        /// <returns>Returns true if the text was found and the replacement successful.</returns>
        public static bool ReplaceNoFormat([NotNull] this Range source, [NotNull] string textToFind, string replacementText)
        {
            return ReplaceNoFormat(source, textToFind, replacementText, resetReplacementFont: false, caseSensitive: true);
        }

        /// <summary>Does a plain-text replacement of text within a Word source.</summary>
        /// <returns>Returns true if the text was found and the replacement successful.</returns>
        public static bool ReplaceNoFormat([NotNull] this Range source, [ItemNotNull] string[] textToFind,
                                           string[] replacementText, bool resetReplacementFont,
                                           bool caseSensitive)
        {
            CheckNotNull(source);
            if (string.IsNullOrEmpty(source.Text))
            {
                return false;
            }

            if (textToFind.Any(string.IsNullOrEmpty))
            {
                throw new ArgumentNullException(nameof(textToFind));
            }

            var retVal = false;
            var refRange = source.Duplicate;
            var refRangeText = refRange.Text;
            var dupRange = refRange.Duplicate;
            var resetRanges = false;

            for (var idx = 0; idx <= textToFind.GetUpperBound(dimension: 0); idx++)
            {
                var findText = textToFind[idx];
                var newText = replacementText[idx];

                if (findText == newText)
                    continue;

                if (resetRanges)
                {
                    refRange.SetRange(source.Start, source.End);
                    refRangeText = refRange.Text;
                }

                var idxFind = refRangeText.IndexOf(findText, caseSensitive ? StringComparison.Ordinal : StringComparison.InvariantCultureIgnoreCase);

                while (idxFind >= 0)
                {
                    var startIdx = refRange.Start + idxFind;
                    dupRange.SetRange(startIdx, startIdx + findText.Length);
                    dupRange.Text = newText;
                    retVal = true;
                    if (resetReplacementFont)
                    {
                        dupRange.Font.Reset();
                    }

                    refRange.Start = dupRange.End;
                    resetRanges = true;
                    idxFind = refRange.Text?.IndexOf(findText, StringComparison.Ordinal) ?? -1;
                }
            }

            Marshal.ReleaseComObject(refRange);
            Marshal.ReleaseComObject(dupRange);

            return retVal;
        }

        /// <summary>Does a plain-text replacement of text within a Word source.</summary>
        /// <param name="source">Input Word source.</param>
        /// <param name="textToFind">Text to replace.</param>
        /// <param name="replacementText">Replacement text.</param>
        /// <param name="resetReplacementFont">If true, then remove all font formatting from the replacement text.</param>
        /// <param name="caseSensitive">If true, does not replace text that does not match the case of textToFind.</param>
        /// <returns>Returns true if the text was found and the replacement successful.</returns>
        public static bool ReplaceNoFormat([NotNull] this Range source, [NotNull] string textToFind, string replacementText,
                                           bool resetReplacementFont, bool caseSensitive)
        {
            CheckNotNull(source);
            if (string.IsNullOrEmpty(source.Text))
            {
                return false;
            }

            Check.NotNull(textToFind, nameof(textToFind));
            if (string.IsNullOrEmpty(textToFind))
            {
                throw new ArgumentNullException(nameof(textToFind));
            }

            var retVal = false;
            var refRange = source.Duplicate;
            var dupRange = refRange.Duplicate;

            var idxFind = refRange.Text.IndexOf(textToFind,
                                                caseSensitive
                                                    ? StringComparison.Ordinal
                                                    : StringComparison.InvariantCultureIgnoreCase);

            while (idxFind >= 0)
            {
                var startIdx = refRange.Start + idxFind;
                dupRange.SetRange(startIdx, startIdx + textToFind.Length);
                dupRange.Text = replacementText;
                retVal = true;
                if (resetReplacementFont)
                {
                    dupRange.Font.Reset();
                }

                refRange.Start = dupRange.End;
                idxFind = refRange.Text == null ? -1 : refRange.Text.IndexOf(textToFind, StringComparison.Ordinal);
            }

            Marshal.ReleaseComObject(refRange);
            Marshal.ReleaseComObject(dupRange);

            return retVal;
        }


        public static string ToXml([NotNull] this Range source)
        {
            CheckNotNull(source);
            throw new NotImplementedException("ToXml");
        }


        public static string ToHtml([NotNull] this Range source)
        {
            CheckNotNull(source);
            throw new NotImplementedException("ToHtml");
        }


        public static string ToCleanText([NotNull] this Range source)
        {
            CheckNotNull(source);
            return source.Text.Clean();
        }


        public static void Clean(this Range source)
        {
            source.Clean(true, true, true, true, true, true, null);
        }

        public static void Clean(this Range source, bool unlinkHyperlinks, bool deleteShapes, bool removeHiddenCharacters, bool normalizeWhiteSpace, bool replaceNonBreakingHyphens, bool acceptRevisions,
                                 string[] replaceTextWithSpaces)
        {
            // Unlink hyperlinks; remove HTML ActiveX fields
            if (unlinkHyperlinks)
            {
                foreach (Field fld in source.Fields)
                    switch (fld.Type)
                    {
                        case WdFieldType.wdFieldHyperlink:
                            fld.Unlink();
                            break;
                        case WdFieldType.wdFieldIndexEntry:
                        case WdFieldType.wdFieldHTMLActiveX:
                            fld.Delete();
                            break;
                    }
            }

            // Delete shapes
            if (deleteShapes)
            {
                foreach (Shape shp in source.ShapeRangeLJ())
                    shp.Delete();
                foreach (InlineShape shp in source.InlineShapes)
                    shp.Delete();
            }

            // Trim hidden characters
            if (removeHiddenCharacters && source.Font.Hidden != 0)
            {
                var showAll = source.Document.ActiveWindow.View.ShowAll;
                source.Document.ActiveWindow.View.ShowAll = true;
                try
                {
                    var dupRange = source.Duplicate;
                    while (dupRange.Start < source.End && source.Font.Hidden != 0)
                    {
                        // Find the first hidden character
                        dupRange.Collapse(WdCollapseDirection.wdCollapseStart);
                        dupRange.MoveEnd(WdUnits.wdCharacter, Count: 1);
                        while (dupRange.Font.Hidden == 0 && dupRange.End < source.End)
                            dupRange.MoveEnd(WdUnits.wdCharacter, Count: 1);
                        if (dupRange.Font.Hidden == 0)
                            continue;

                        dupRange.Start = dupRange.End - 1;
                        while (dupRange.Font.Hidden == -1 && dupRange.End < source.End)
                            dupRange.End += 1;
                        if (dupRange.Font.Hidden == (int)WdConstants.wdUndefined)
                            dupRange.End -= 1;
                        dupRange.Text = string.Empty;
                        dupRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    }
                }
                finally
                {
                    source.Document.ActiveWindow.View.ShowAll = showAll;
                }
            }

            // Trim white space characters
            if (normalizeWhiteSpace && !string.IsNullOrEmpty(source.Text))
            {
                // Convert white space and replacement char arrays to string arrays
                var whiteSpaceArray = string.Empty.WhiteSpaceCharacters().Select(c => c.ToString()).ToArray();
                var spaceArray = new string[whiteSpaceArray.Length].Populate(" ");

                source.ReplaceNoFormat(whiteSpaceArray, spaceArray, resetReplacementFont: false, caseSensitive: false);
            }

            // Normalize non-breaking hyphens
            if (replaceNonBreakingHyphens)
            {
                source.Replace(Convert.ToChar(30).ToString(), "-",
                               matchCase: false,
                               matchWholeWord: false,
                               validateRange: false,
                               maxCountForValidate: 3,
                               resetReplacementFont: false);
            }

            if (replaceTextWithSpaces != null)
            {
                foreach (var textToDelete in replaceTextWithSpaces)
                {
                    source.ReplaceNoFormat(textToDelete, " ");
                }
            }

            // Accept revisions
            if (acceptRevisions)
            {
                var text = source.Text;
                var textLength = text?.Length ?? 0;
                if (textLength != source.End - source.Start)
                {
                    foreach (var rev in source.RevisionsLJ().Reverse())
                        rev.Accept();
                }
            }
        }


        /// <summary>Gets the hyperlink field that contains a Word range.</summary>
        /// <returns>Returns the containing hyperlink field, or null if none exists.</returns>
        [CLSCompliant(false)]
        public static Field HyperlinkResultContainer([NotNull] this Range source)
        {
            CheckNotNull(source);

            var fields = source.FieldContainers();

            // ReSharper disable once LoopCanBeConvertedToQuery
            foreach (var fld in fields)
            {
                if (fld.ResultLJ() == null) continue;
                if (source.Start >= fld.Result.Start && source.End <= fld.Result.End)
                {
                    return fld;
                }
            }

            return null;
        }


        /// <summary>Determine the number of characters in a range.</summary>
        /// <returns>The number of characters in the range, not including any deleted text.</returns>
        public static int CharacterCount([NotNull] this Range source)
        {
            CheckNotNull(source);
            return source.CharacterCount(includeDeletedText: false);
        }

        /// <summary>Determine the number of characters in a range.</summary>
        /// <returns>The number of characters in the range.</returns>
        public static int CharacterCount([NotNull] this Range source, bool includeDeletedText)
        {
            CheckNotNull(source);

            // Range.Characters includes deleted text
            var charCount = source.Characters.Count;
            if (includeDeletedText) return charCount;

            // If not including deleted text, subtract out the number of deleted characters.
            foreach (var rev in source.RevisionsLJ())
            {
                if (rev.Type != WdRevisionType.wdRevisionDelete) continue;

                var revDup = rev.Range.Duplicate;
                revDup.Start = Math.Max(source.Start, revDup.Start);
                revDup.End = Math.Min(source.End, revDup.End);
                charCount -= string.IsNullOrEmpty(revDup.Text) ? 0 : revDup.Text.Length;
            }

            return charCount;
        }


        /// <summary>Gets the width of the section text, in points</summary>
        /// <param name="source">A Range instance.</param>
        /// <param name="checkOtherSectionsInOrder">
        ///     If true, evaluates sections in document order. The first section without an
        ///     undefined text width is used.
        /// </param>
        /// <param name="defaultTextWidth">If the text width cannot be found, sets the value returned.</param>
        /// <returns>
        ///     The width of the text (in points). If any of the PageWidth,
        ///     LeftMargin, RightMargin, or Gutter values are undefined, returns WdConstants.wdUndefined (9999999F).
        /// </returns>
        // ReSharper disable once InconsistentNaming
        public static float TextWidthLJ([NotNull] this Range source, bool checkOtherSectionsInOrder = true, float? defaultTextWidth = null)
        {
            const float undefinedValue = (float)WdConstants.wdUndefined;

            var sct = source.Sections.First;
            var textWidth = sct.TextWidthLJ();
            if (textWidth < undefinedValue) // != wdUndefined
                return textWidth;

            if (checkOtherSectionsInOrder)
            {
                // At least one of the values is wdUndefined
                // Try getting a value from another section; test in order
                var doc = source.Document;
                for (var idx = 1; idx <= doc.Sections.Count; idx++)
                {
                    if (idx == sct.Index)
                        continue;

                    var altSct = doc.Sections[idx];
                    textWidth = altSct.TextWidthLJ();
                    if (textWidth < undefinedValue) // != wdUndefined
                        return textWidth;
                }
            }

            return defaultTextWidth != null ? sct.Application.InchesToPoints(defaultTextWidth.GetValueOrDefault()) : undefinedValue;
        }


        /// <summary>Determines if a Word source is in a Word content control.</summary>
        /// <param name="source">Word source to analyze.</param>
        /// <returns>Returns true if source is in a content control.</returns>
        public static bool InContentControl([NotNull] this Range source)
        {
            CheckNotNull(source);
            if (source.Application.VersionLJ() <= OfficeVersion.Office2007)
            {
                return false;
            }

            return ContentControlContainer(source) != null;
        }


        /// <summary>Determines if source comes just before a content control.</summary>
        /// <returns>Returns true if source end is at the start of a content control.</returns>
        public static bool PreceedsContentControl([NotNull] this Range source)
        {
            CheckNotNull(source);
            var rng = source.Document.Range();
            var result = rng.ContentControlRangesLJ().Any(ccRange => ccRange.Start == source.End);
            Marshal.ReleaseComObject(rng);
            return result;
        }


        /// <summary>Determine if source starts just after a content control.</summary>
        /// <returns>Returns true if source start is just after the end of a content control.</returns>
        public static bool FollowsContentControl([NotNull] this Range source)
        {
            CheckNotNull(source);
            var rng = source.Document.Range();
            var result = rng.ContentControlRangesLJ().Any(ccRange => source.Start == ccRange.End + 1);
            Marshal.ReleaseComObject(rng);
            return result;
        }


        /// <summary>Finds the content control containing the source.</summary>
        /// <returns>Returns containing content control.</returns>
        public static ContentControl ContentControlContainer([NotNull] this Range source)
        {
            Check.NotNull(source, nameof(source));
            if (source.Application.VersionLJ() <= OfficeVersion.Office2007)
                return null;

            return source.Document.ContentControls.Cast<ContentControl>().FirstOrDefault(cc => cc.Range.Start >= source.Start && cc.Range.End >= source.End);
        }


        /// <summary>Gets the tag of the containing content control. Returns empty string if there isn't one.</summary>
        public static string ContentContainerTag([NotNull] this Range source)
        {
            Check.NotNull(source, nameof(source));
            var cc = ContentControlContainer(source);
            return cc != null ? cc.Tag : string.Empty;
        }


        /// <summary>Determines if source is in a MacPac content control by analyzing the content control tag.</summary>
        public static bool InMacPacContentControl([NotNull] this Range source)
        {
            return source.InMacPacContentControl(new[] { "mps", "mpb" });
        }


        public static bool InMacPacContentControl([NotNull] this Range source, string[] prefixes)
        {
            Check.NotNull(source, nameof(source));
            var tagText = ContentContainerTag(source);
            return !string.IsNullOrEmpty(tagText) &&
                   prefixes.Any(prefix => tagText.ToLower().StartsWith(prefix.ToLower()));
        }


        /// <summary>Gets all ContentControls fully contained within the source range.</summary>
        /// <param name="source">A Word.Range object.</param>
        /// <returns>Returns an IEnumerable (of ContentControl)</returns>
        [CLSCompliant(false)]
        public static IEnumerable<ContentControl> ContainedContentControls([NotNull] this Range source)
        {
            // ContentControls not supported before Word 2007
            if (source.Application.VersionLJ() < OfficeVersion.Office2007)
                return Enumerable.Empty<ContentControl>();

            var ccs = source.Document.ContentControls;
            return from ContentControl cc in ccs where cc.Range.InRange(source) select cc;
        }


        /// <summary>Determines whether a specified ContentControl is contained within the source range.</summary>
        /// <param name="source">A Word.Range object.</param>
        /// <param name="contentControl">A Word.ContentControl object.</param>
        /// <returns>Returns true if the ContentControl object is contained within the source range.</returns>
        [CLSCompliant(false)]
        public static bool ContainsContentControl([NotNull] this Range source, [NotNull] ContentControl contentControl)
        {
            // ContentControls not supported before Word 2007
            return source.Application.VersionLJ() >= OfficeVersion.Office2007 && contentControl.Range.InRange(source);
        }


        /// <summary>Deletes all content controls in the source.</summary>
        public static void DeleteContentControls([NotNull] this Range source)
        {
            Check.NotNull(source, nameof(source));

            // A ContentControl may contain content controls.
            // If the parent gets deleted prior to one or more of its contained controls,
            // the contained controls will already be deleted in the loop.
            // Ignore this error.

            var ccs = source.ContainedContentControls().ToList();
            foreach (var cc in ccs)
            {
                try
                {
                    var lockCc = cc.LockContentControl;

                    // Will not get here if the CC is deleted
                    if (lockCc)
                        cc.LockContentControl = false;
                    if (cc.LockContents)
                        cc.LockContents = false;
                    cc.Delete();
                }
                catch (Exception ex)
                {
                    // Ignore error
                }
            }
        }

        /// <summary>Deletes all ContentControls in the document.</summary>
        /// <param name="source">A Word document.</param>
        [CLSCompliant(false)]
        public static void DeleteContentControls([NotNull] this Document source)
        {
            // ContentControls not supported before Word 2007
            if (source.Application.VersionLJ() < OfficeVersion.Office2007)
                return;

            foreach (Range storyRange in source.StoryRanges)
            {
                storyRange.DeleteContentControls();
            }
        }


        /// <summary>Deletes all content controls in the document.</summary>
        public static void DeleteContentControls([NotNull] this _Document wordDocument)
        {
            wordDocument.DeleteContentControls(wordDocument.StoryRanges.Cast<Range>().Select(storyRange => storyRange));
        }

        /// <summary>Deletes all content controls in input story ranges in the document.</summary>
        public static void DeleteContentControls([NotNull] this _Document wordDocument, IEnumerable<Range> storyRanges)
        {
            foreach (var storyRange in storyRanges)
            {
                storyRange.DeleteContentControls();
            }
        }


        public static void DeleteEmptyParagraphs([NotNull] this Range source, bool alsoDeleteWhitespaceParagraphs = false)
        {
            var paras = source.Paragraphs.Cast<Paragraph>().ToList();
            foreach (var para in paras)
            {
                var paraRange = para.Range;
                var paraText = paraRange.Text;
                var deletePara = alsoDeleteWhitespaceParagraphs
                                     ? string.IsNullOrWhiteSpace(paraText)
                                     : string.IsNullOrEmpty(paraText);

                if (deletePara)
                    paraRange.Delete();
            }
        }


        [CLSCompliant(false)]
        public static void AppendRange(this Document source, Range rangeToAppend, bool insideFieldCode, bool newParagraph = true)
        {
            Check.NotNull(source, nameof(source));
            Check.NotNull(rangeToAppend, nameof(rangeToAppend));

            // Copy in formatted text
            var workDocRange = source.Range(0, 0);
            workDocRange.Move(WdUnits.wdStory); // Moves to the end of the story
            workDocRange.AppendRange(rangeToAppend, insideFieldCode, newParagraph);
        }

        /// <summary>Appends the formatted text of a Word range to the end of a range.</summary>
        /// <param name="source">Location in the document to place the new formatted text.</param>
        /// <param name="fromRange">Source range of the formatted text.</param>
        /// <param name="insideFieldCode">If true, fromRange is inside a field code (for getting safe subranges).</param>
        /// <param name="newParagraph">If true, inserts paragraph return between source and fromRange.</param>
        public static void AppendRange([NotNull] this Range source, [NotNull] Range fromRange, bool insideFieldCode, bool newParagraph)
        {
            Check.NotNull(source, nameof(source));
            Check.NotNull(fromRange, nameof(fromRange));

            var startPos = source.Start;

            var workRange = source.Duplicate;
            workRange.Collapse(WdCollapseDirection.wdCollapseEnd);

            if (newParagraph)
            {
                workRange.Text = ControlStringCharacters.CharacterReturn;
                workRange.Collapse(WdCollapseDirection.wdCollapseStart);
            }
            else
            {
                var paras = workRange.Paragraphs;
                var firstPara = paras[1];
                var firstParaRange = firstPara.Range;
                if (workRange.Start == firstParaRange.Start)
                {
                    workRange.Move(WdUnits.wdCharacter, -1);
                }

                Marshal.ReleaseComObject(firstParaRange);
                Marshal.ReleaseComObject(firstPara);
                Marshal.ReleaseComObject(paras);
            }

            //// Copy in formatted text

            foreach (var rng in fromRange.SafeSubRanges(insideFieldCode))
            {
                workRange.FormattedText = rng.FormattedText;
                workRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                Marshal.ReleaseComObject(rng);
            }

            // Reset source range to include entire fromRange
            source.SetRange(startPos, source.End);
        }


        /// <summary>Determines if a Word source contains parts of fields, revisions, or table cells (i.e., "well-formed").</summary>
        /// <returns>Returns true if source is well-formed.</returns>
        public static bool IsSafe([NotNull] this Range source)
        {
            CheckNotNull(source);

            var allTxt = source.Text;
            if (allTxt == null)
            {
                // It's only AllSafe if source is collapsed
                return source.End == source.Start;
            }

            var hasCellMarks = allTxt.Contains(ControlStringCharacters.Bell);
            if (source.End - source.Start == allTxt.Length && hasCellMarks == false)
            {
                // It's well behaved
                // It could possibly be in a field result, but if that's the case, then it's still safe
                // Note that the last character of a table has a source length of 1 but a text of chr(13)chr(7);
                // that's the only case in which textlen > rangelen,
                // which could hide another iteme where the reverse was true
                return true;
            }

            if (source.Application.VersionLJ() == OfficeVersion.Office2013)
            {
                // Can't check revisions in 2013 because
                // it causes Selection to be collapsed or hidden
                return false;
            }

            return source.Fields.Count == 0 && source.Revisions.Count == 0 && hasCellMarks == false;
        }


        /// <summary>Creates a list of well-formed Word ranges from a Word source.</summary>
        /// <param name="source">Input Word source to sub-divide.</param>
        /// <returns>A WordRangeList of subranges.</returns>
        public static IEnumerable<Range> SafeSubRanges([NotNull] this Range source)
        {
            return SafeSubRanges(source, insideFieldCode: false);
        }

        /// <summary>Creates a list of well-formed Word ranges from a Word source.</summary>
        /// <param name="source">Input Word source to sub-divide.</param>
        /// <param name="insideFieldCode">If true, input Word source is inside a field code.</param>
        /// <returns>A collection of subranges.</returns>
        public static IEnumerable<Range> SafeSubRanges([NotNull] this Range source, bool insideFieldCode)
        {
            CheckNotNull(source);

            //var retList = new WordRangeList();
            var foundItems = false;
            // If there are no fields to worry about,
            // then just return SourceRange as a singleton.
            var rngAll = source.Duplicate;
            rngAll.TextRetrievalMode.IncludeFieldCodes = insideFieldCode;
            rngAll.TextRetrievalMode.IncludeHiddenText = false;

            if (rngAll.IsSafe())
            {
                // Note that the above call has special code for 2013
                yield return rngAll;
                yield break;
            }

            // Strategy:
            //   Fill the array with one or more rngWork objects
            //   Work our way through the Characters collection
            //   Each time that it moves by one position only, we know that we can extend the prior source
            //   But if it jumps, we cut the prior source off, save it, and start a new one

            var rngWork = rngAll.Duplicate;
            rngWork.Collapse(WdCollapseDirection.wdCollapseStart);

            var allCharacters = rngAll.Characters;
            var lPrevEnd = -1;

            foreach (Range rngC in allCharacters)
            {
                var rcTrue = rngC.TrueCharacterRange();

                if (rcTrue.Start > lPrevEnd)
                {
                    // Put the prior rngWork (unless it's collapsed)
                    // And create a new rngWork object
                    if (rngWork.End > rngWork.Start)
                    {
                        //retList.Add(rngWork);
                        foundItems = true;
                        yield return rngWork;
                    }

                    rngWork = rcTrue.Duplicate;
                }
                else
                {
                    rngWork.End = rcTrue.End;
                }

                lPrevEnd = rcTrue.End;

                if (rcTrue != rngC)
                {
                    Marshal.ReleaseComObject(rcTrue);
                }

                Marshal.ReleaseComObject(rngC);
            }

            Marshal.ReleaseComObject(allCharacters);

            if (rngWork.End > rngAll.End)
            {
                // This happens if rngsource ends just before
                // a result, in which case it jumps over end-field position
                rngWork.End = rngAll.End;
            }

            // Put any final source
            if (rngWork.End > rngWork.Start)
            {
                //retList.Add(rngWork);
                foundItems = true;
                yield return rngWork;
            }

            if (foundItems || insideFieldCode)
                yield break;

            // I think this could only happen if it was unexpectedly all inside a field code
            // Or, if the entire source was hidden
            // If so, then we return collapsed source at start of 1st field result
            try
            {
                // Remember, not all fields have results (e.g. TC)
                rngWork = rngAll.Fields[Index: 1].Result.Duplicate;
                rngWork.Collapse(WdCollapseDirection.wdCollapseStart);
            }
            catch (Exception)
            {
                rngWork = null;
            }

            if (rngWork == null)
            {
                try
                {
                    rngWork = rngAll.Fields[Index: 1].Code.Duplicate;
                    rngWork.Collapse(WdCollapseDirection.wdCollapseEnd);
                    // rngwork is now inside the field code, just before the end-of-field marker
                    // next statement puts it just before the next real character
                    rngWork.MoveStart(WdUnits.wdCharacter, -1);
                }
                catch (Exception)
                {
                    rngWork = null;
                }

                if (rngWork == null)
                {
                    rngWork = rngAll.Duplicate;
                    rngWork.Collapse(WdCollapseDirection.wdCollapseEnd);
                }
            }

            yield return rngWork;
        }

        /// <summary>Gets the "true" source of a character in a document.</summary>
        /// <param name="source">Word source containing the character.</param>
        /// <returns>Returns the source of the character.</returns>
        public static Range TrueCharacterRange([NotNull] this Range source)
        {
            CheckNotNull(source);
            if (source.End - source.Start == 1)
                return source;

            var trueRange = source.Duplicate;
            // See if real character is last in source
            // This will happen on the 1st character of a field with a result len >1
            trueRange.Start = trueRange.End - 1;
            if (!string.IsNullOrEmpty(trueRange.Text))
                return trueRange;

            // See if real character is first in source
            // This will happen on the last character of a field with a result len >1
            trueRange.SetRange(source.Start, source.Start + 1);
            if (!string.IsNullOrEmpty(trueRange.Text))
                return trueRange;

            // This can happen on a 1-char result, where
            // there's a pad at each end
            trueRange.SetRange(source.Start, source.End);
            while (!string.IsNullOrEmpty(trueRange.Text))
            {
                trueRange.End -= 1;
            }

            // Were' now just before the actual character
            trueRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            trueRange.End = trueRange.End + 1;
            if (string.IsNullOrEmpty(trueRange.Text))
            {
                // This shouldn't happen!
                // But if it does, then return an empty source
                // rather than risk returning a mal-formed source
                //Debug.Assert(false);
                trueRange.SetRange(source.End, source.End);
            }
            // Debug.Print source.Start, source.End, rCActual.Start, rCActual.End, source, rCActual

            return trueRange;
        }


        /// <summary>
        ///     Determines if the plain text representation of the text of a Word source matches that of a RichTextBoxTom
        ///     TextRange.
        /// </summary>
        /// <param name="source">Word source to analyze.</param>
        /// <param name="textRange">TextRange to analyze.</param>
        /// <param name="caseSensitive">If true, case must match.</param>
        /// <param name="checkForMatchingParens">
        ///     If true, input word Range should have same number of open and close parentheses. This
        ///     check is necessary because the Word API can represent a symbol incorrectly as an open parenthesis.
        /// </param>
        /// <returns>Returns true if the cleaned text representations are equal.</returns>
        public static bool CleanedTextMatches([NotNull] this Range source, [NotNull] TextRange textRange, bool caseSensitive, bool checkForMatchingParens)
        {
            CheckNotNull(source);
            Check.NotNull(textRange, nameof(textRange));
            return CleanedTextMatches(source, textRange.Text, caseSensitive, checkForMatchingParens);
        }


        public static bool CleanedTextMatches([NotNull] this Range source, [CanBeNull] string otherText, bool caseSensitive, bool checkForMatchingParens)
        {
            CheckNotNull(source);
            if (string.IsNullOrEmpty(otherText))
            {
                // Return true if the source is collapsed
                return source.Start == source.End;
            }

            var wordText = source.Text.Clean(true);
            var testText = otherText.Clean(true);
            if (!caseSensitive)
            {
                wordText = wordText.ToLower();
                testText = testText.ToLower();
            }

            if (wordText == testText)
            {
                return true;
            }

            // SPECIAL CASE: Symbol characters don't always get read correctly by the Word API.
            if (checkForMatchingParens && wordText.CountOf("(") > wordText.CountOf(")"))
            {
                // If the wordText has more OpenParen than CloseParen, it's likely that difference is due to symbol chars
                // which render in text as "(", so if all other characters match , we'll treat as the same.
                return wordText.Replace('(', ' ') == testText.Replace('(', ' ');
            }

            return false;
        }


        public static string ToEmphasisText([NotNull] this Range source)
        {
            CheckNotNull(source);
            return source.ToEmphasisText(false);
        }

        public static string ToEmphasisText([NotNull] this Range source, bool analyzeByCharacter)
        {
            CheckNotNull(source);

            if (source.Length() == 1)
            {
                if (IsSpecialChar(source.Text))
                    analyzeByCharacter = true;
            }

            var sb2 = new StringBuilder();
            IEnumerable<Range> ranges;

            if (!analyzeByCharacter)
                ranges = SafeSubRanges(source);
            else
            {
                ranges = CharacterRanges(source);
            }

            foreach (var rng in ranges)
            {
                EmphasisStringFromRangeCore(rng, sb2);
                Marshal.ReleaseComObject(rng);
            }

            //rngList.ForEach(rng => EmphasisStringFromRangeCore(rng, sb2));
            //rngList.Dispose();

            return sb2.ToString();

            bool IsSpecialChar(string cTest) => cTest == ControlStringCharacters.StartOfText || cTest == ControlStringCharacters.StartOfHeading;

            IEnumerable<Range> CharacterRanges(Range src)
            {
                foreach (Range c in src.Characters)
                {
                    if (!IsSpecialChar(c.Text))
                        yield return c;
                    else
                        Marshal.ReleaseComObject(c);
                }
            }
        }

        private static void EmphasisStringFromRangeCore(Range wordRange, StringBuilder sb)
        {
            // Note overload (below)

            const int undefined = (int)WdConstants.wdUndefined;

            var spanRange = wordRange.Duplicate;
            var spanFont = spanRange.Font;
            var spanFind = spanRange.Find;
            var spanFindFont = spanFind.Font;
            var remainderRange = wordRange.Duplicate;
            var remainderFont = remainderRange.Font;
            var remainderItalic = Convert.ToInt32(undefined);
            var remainderUnderline = Convert.ToInt32(undefined);
            var remainderSmallCaps = Convert.ToInt32(undefined);

            spanRange.Collapse(WdCollapseDirection.wdCollapseStart);
            spanFind.ClearAllFuzzyOptions();

            var limit = wordRange.End;
            var priorStart = -1;

            while (remainderRange.Start < limit && remainderRange.Start > priorStart) // if you reach the end of the document, start won't increment
            {
                priorStart = remainderRange.Start;

                // Once we've got a definite value, no reason to re-query word
                if (remainderItalic == undefined)
                {
                    remainderItalic = remainderFont.Italic;
                }

                if (remainderUnderline == undefined)
                {
                    remainderUnderline = (int)remainderFont.Underline;
                }

                if (remainderSmallCaps == undefined)
                {
                    remainderSmallCaps = remainderFont.SmallCaps;
                }

                // Get the values while it's a collapsed range
                //	Use known values from remainder where possible
                var emph = CharacterFormattingEmphasis.None;
                if (remainderItalic == undefined)
                {
                    if (spanFont.Italic != 0)
                    {
                        emph = CharacterFormattingEmphasis.Italic;
                    }
                }
                else
                {
                    if (remainderItalic != 0)
                    {
                        emph = CharacterFormattingEmphasis.Italic;
                    }
                }

                if (remainderUnderline == undefined)
                {
                    if (spanFont.Underline != 0)
                    {
                        emph |= CharacterFormattingEmphasis.Underline;
                    }
                }
                else
                {
                    if (remainderUnderline != 0)
                    {
                        emph |= CharacterFormattingEmphasis.Underline;
                    }
                }

                if (remainderSmallCaps == undefined)
                {
                    if (spanFont.SmallCaps != 0)
                    {
                        emph |= CharacterFormattingEmphasis.SmallCaps;
                    }
                }
                else
                {
                    if (remainderSmallCaps != 0)
                    {
                        emph |= CharacterFormattingEmphasis.SmallCaps;
                    }
                }


                if (remainderItalic != undefined && remainderUnderline != undefined && remainderSmallCaps != undefined)
                {
                    // Remainder is uniform, so no need to search
                    spanRange.End = limit;
                }
                else
                {
                    //	We will search on heterogeneous properties to find a uniform string
                    spanFind.ClearFormatting();
                    if (remainderItalic == undefined)
                    {
                        spanFindFont.Italic = spanFont.Italic;
                    }

                    if (remainderUnderline == undefined)
                    {
                        spanFindFont.Underline = spanFont.Underline;
                    }

                    if (remainderSmallCaps == undefined)
                    {
                        spanFindFont.SmallCaps = spanFont.SmallCaps;
                    }

                    // Remainder is heterogeneous, so Find should never go past end
                    spanFind.Execute();
                    if (spanRange.End > limit)
                    {
                        // but just in case it does go past end
                        spanRange.End = limit;
                    }
                }

                var chars = spanRange.Characters;
                sb.Append(Convert.ToInt16(emph).ToString("X")[0], chars.Count); // "X" means Hex
                Marshal.ReleaseComObject(chars);

                spanRange.Collapse(WdCollapseDirection.wdCollapseEnd);

                remainderRange.Start = spanRange.Start;
            }

            Marshal.ReleaseComObject(spanRange);
            Marshal.ReleaseComObject(spanFont);
            Marshal.ReleaseComObject(spanFind);
            Marshal.ReleaseComObject(spanFindFont);
            Marshal.ReleaseComObject(remainderRange);
            Marshal.ReleaseComObject(remainderFont);
        }


        public static void ConvertToSmartQuotes([NotNull] this Range source, bool includeFootEndNotes = false)
        {
            CheckNotNull(source);

            // Use the Word.Range.Find to find the dumb quotes and
            // re-format them as smart quotes. NOTE: in the case where
            // the user has entered in two single quotes instead of
            // one double quote, run the loops twice... this gets around
            // a problem in the Word API that refuses to AutoFormat
            // the first single quote if it is followed by another.
            for (var rep = 1; rep <= 2; rep++)
            {
                foreach (var asciiChar in string.Empty.PlainQuoteCharacters())
                {
                    var charText = asciiChar.ToString();

                    var lastStart = -1;
                    var endOfRange = source.End;
                    var workRange = source;
                    workRange.Collapse(WdCollapseDirection.wdCollapseStart);
                    bool done;
                    var find = workRange.Find;
                    find.ClearFormatting();
                    find.Text = charText;
                    find.Forward = true;
                    find.Format = false;
                    find.Wrap = WdFindWrap.wdFindStop;
                    find.Execute();

                    do
                    {
                        done = !find.Found || workRange.End > endOfRange || workRange.Start <= lastStart;
                        if (done) continue;

                        lastStart = workRange.Start;
                        if (workRange.InRange(source) && workRange.Text == charText)
                            workRange.AutoFormat(); // Does the conversion to smart quotes (most of the time, see above)
                        workRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                        find.Execute();
                    } while (!done);

                    Marshal.ReleaseComObject(find);
                } // quoteChar
            }

            if (includeFootEndNotes)
            {
                foreach (Footnote fn in source.Footnotes)
                    fn.Range.ConvertToSmartQuotes();
                foreach (Endnote en in source.Endnotes)
                    en.Range.ConvertToSmartQuotes();
            }
        }


        public static void MoveBoundaries([NotNull] this Range source, string characters, string expandOrCollapse, string startEndOrBoth)
        {
            MoveBoundaries(source, characters.ToCharArray(), expandOrCollapse, startEndOrBoth);
        }

        public static void MoveBoundaries([NotNull] this Range source, char[] characters, string expandOrCollapse, string startEndOrBoth)
        {
            startEndOrBoth = startEndOrBoth.ToLower();
            var doStart = startEndOrBoth != "end";
            var doEnd = startEndOrBoth != "start";
            var expand = expandOrCollapse.ToLower() == "expand";

            var workRange = source.Duplicate;
            try
            {
                if (doStart)
                {
                    var done = false;
                    while (!done)
                    {
                        done = InchwormWorkRange(expand) == 0 || string.IsNullOrEmpty(workRange.Text);
                        if (!done)
                        {
                            if (!characters.Contains(workRange.Text[0]))
                            {
                                done = true;
                                source.Start = expand ? workRange.End : workRange.Start;
                            }
                        }
                    }
                }

                if (doEnd)
                {
                    var done = false;
                    while (!done)
                    {
                        done = InchwormWorkRange(!expand) == 0 || string.IsNullOrEmpty(workRange.Text);
                        if (!done)
                        {
                            if (!characters.Contains(workRange.Text[0]))
                            {
                                done = true;
                                source.End = expand ? workRange.Start : workRange.End;
                            }
                        }
                    }
                }
            }
            finally
            {
                if (workRange != null)
                    Marshal.ReleaseComObject(workRange);
            }

            int InchwormWorkRange(bool inchwormBackwards)
            {
                // Inchworm Scenarios
                // 1. Expanding at the start; Inchworm backwards
                // 2. Expanding at the end; Inchworm forwards
                // 3. Collapsing at the start; Inchworm forwards
                // 4. Collapsing at the end; Inchworm backwards
                if (inchwormBackwards)
                {
                    workRange.Collapse(WdCollapseDirection.wdCollapseStart);
                    return workRange.MoveStart(WdUnits.wdCharacter, -1);
                }

                // Inchworm forwards
                workRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                return workRange.MoveEnd(WdUnits.wdCharacter, 1);
            }
        }
    }
}