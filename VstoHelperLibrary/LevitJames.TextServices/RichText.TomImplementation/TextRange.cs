// © Copyright 2018 Levit & James, Inc.
//#define TRACK_DISPOSED

using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

//*** used when TRACK_DISPOSED is enabled***

namespace LevitJames.TextServices
{
#pragma warning disable CS0659 // Type overrides Object.Equals(object o) but does not override Object.GetHashCode()


    /// <summary>
    /// TextRange objects are powerful editing and data-binding tools that allow a program to select text in a story and then examine or change that text.
    /// </summary>
    /// <remarks>Open a web browser browser and search for ITextRange.xxxx for detailed documentation.</remarks>
    public class TextRange : ITextRange, IEquatable<TextRange>, IDisposable
#pragma warning restore CS0659 // Type overrides Object.Equals(object o) but does not override Object.GetHashCode()
    {
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private ITextRange _range;

#if (TRACK_DISPOSED)
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private readonly string _disposedSource;
#endif

        internal ITextRange Range => _range;

        internal TextRange(ITextRange range)
        {
            _range = range;
#if (TRACK_DISPOSED)
            _disposedSource = LevitJames.Core.Diagnostics.DisposeTracker.GetSource();
#endif
        }

        [DebuggerHidden]
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        // ReSharper disable once InconsistentNaming
        internal ITextRange ITextRange => _range;


        /// <summary>
        ///     Gets or sets the plain text in this range.
        /// </summary>


        /// <remarks>
        ///     The text returned by the Text property is given in Unicode.
        ///     The end-of-paragraph mark may be given by 0x2029 (the Unicode Paragraph Separator),
        ///     or by carriage return/line feed (CRLF) (0xd, 0xa), or by CR alone, depending on the original file.
        ///     Microsoft Word uses CR alone, unless it reads another choice in from a file,  the Clipboard, or an IDataObject.
        ///     The placeholder for an embedded object is given by the special character,
        ///     WCH_EMBEDDING, which has the official Unicode value 0xFFFC
        /// </remarks>
        public string Text
        {
            get => _range.Text;
            set => _range.Text = value;
        }

        /// <summary>
        ///     Gets or sets the character at the range's start position.
        /// </summary>


        /// <remarks>
        ///     Similarly, setting TextRange.Char overwrites the first character with the character,
        ///     Char. Note that the characters retrieved and set by these methods are int variables, which hide the way that they
        ///     are stored in the
        ///     backing store (as bytes, words, variable-length, and so forth), and they do not require using a Unicode
        ///     System.String.
        ///     The Char property, which can do most things that a characters collection can, has two big advantages:
        ///     It can reference any character in the parent story instead of being limited to the parent range.
        ///     It is significantly faster, since System.Char are involved instead of range objects.
        ///     Accordingly, the Text Object Model (TOM) does not support a characters collection.
        /// </remarks>
        public int Char
        {
            get => _range.Char;
            set => _range.Char = value;
        }

        /// <summary>
        ///     Creates a duplicate TextRange object.
        /// </summary>


        /// <remarks>
        ///     To create an insertion point in order to traverse a range, first duplicate the range and then collapse the
        ///     duplicate at its Start cp. Note,
        ///     a range is characterized by cpFirst, cpLim, and the story it belongs to.
        ///     Even if the range is actually an TextSelection, the duplicate returned is an TextRange. For an example, see the
        ///     TextRange.FindText method.
        /// </remarks>
        /// )]
        public TextRange Duplicate => new TextRange(_range.Duplicate);

        [DebuggerHidden]
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        ITextRange ITextRange.Duplicate => _range.Duplicate;

        /// <summary>
        /// </summary>

        public string text;

        public TextRange FormattedText
        {
            get => this;
            set => _range.FormattedText = value._range;
        }

        /// <summary>
        /// Copies formatted text (text and its formatting) from the provided source range to this range.
        /// </summary>
        /// <param name="source">The source range to copy from.</param>
        /// <param name="disposeSource">true to dispose of the source range before returning.</param>
        public void CopyFormattedText(TextRange source, bool disposeSource = false)
        {
            var ft = source._range.FormattedText;
            _range.FormattedText = ft;
            Marshal.ReleaseComObject(ft);
            if (disposeSource)
                source.Dispose();
        }

        /// <summary>
        ///     Copies the Paragraph and Font properties from the source range to this TextRange
        /// </summary>
        /// <param name="source">
        ///     The Text Range to copy the Paragraph and Font properties from. It is recommended that the Range
        ///     formatting is continuous to avoid any undefined properties.
        /// </param>
        /// <param name="disposeSource">true to dispose of the source range after the copy.</param>
        public void CopyFormatting(TextRange source, bool disposeSource = false)
        {
            var para = source._range.Para;
            var font = source._range.Font;

            var _ = para.RightIndent;
            var __ = font.Size;
            _range.Para = para;
            _range.Font = font;
            //if (para.TabCount > 0)
            //{
            //    var destPara = _range.Para;
            //    destPara.ClearAllTabs();
            //    for (var i = 0; i < para.TabCount; i++)
            //    {
            //        para.GetTab((ParagraphTabIndex)i, out var fTabPos, out var tabAlign, out var labLeader);
            //        destPara.AddTab(fTabPos, tabAlign, labLeader);
            //    }
            //    Marshal.ReleaseComObject(destPara);
            //}

            Marshal.ReleaseComObject(para);
            Marshal.ReleaseComObject(font);

            if (disposeSource)
                source.Dispose();
        }

        [DebuggerHidden]
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        ITextRange ITextRange.FormattedText
        {
            get => _range.FormattedText;
            set => _range.FormattedText = value;
        }

        /// <summary>
        ///     Gets or sets the specified range Start.
        /// </summary>

        /// <returns>The start position of the range object.</returns>

        public int Start
        {
            get => _range.Start;
            set => _range.Start = value;
        }

        /// <summary>
        ///     Gets or sets the specified range End.
        /// </summary>

        /// <returns>The end position of the range object.</returns>

        public int End
        {
            get => _range.End;
            set => _range.End = value;
        }

        /// <summary>
        ///     Returns the TextFont object with the character attributes of the specified range.
        /// </summary>

        /// <returns>A TextFont instance.</returns>
        /// <remarks>
        ///     For plain-text controls, these objects do not vary from range to range,
        ///     but in rich-text solutions, they do. See the section on TextFont for further details.
        /// </remarks>
        public TextFont Font
        {
            get => new TextFont(_range.Font);
            set => _range.Font = value;
        }

        [DebuggerHidden]
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        ITextFont ITextRange.Font
        {
            get => _range.Font;
            set => _range.Font = value;
        }

        /// <summary>
        ///     Returns an TextParagraph object with the paragraph attributes of the specified range
        /// </summary>

        /// <returns>A TextParagraph instance.</returns>

        public TextParagraph Paragraph
        {
            get => new TextParagraph(_range.Para);
            set => _range.Para = value;
        }

        [DebuggerHidden]
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        ITextPara ITextRange.Para
        {
            get => _range.Para;
            set => _range.Para = value;
        }

        /// <summary>
        ///     Returns the count of characters in the specified range's story.
        /// </summary>

        /// <returns>A count of the characters in the specified range's story</returns>

        public int StoryLength => _range.StoryLength;

        /// <summary>
        ///     Returns the type of the specified range's story.
        /// </summary>

        /// <returns>One of the values defined in RangeStoryType.</returns>
        /// <remarks>Currently the RichEdit control only supports RangeStoryType.MainText</remarks>
        public RangeStoryType StoryType => _range.StoryType;

        /// <summary>
        /// Adjusts the range endpoints to the specified values.
        /// </summary>
        /// <param name="active"></param>
        /// <param name="other"></param>
        public void SetRange(int active, int other)
        {
            _range.SetRange(active, other);
        }

        /// <summary>
        /// Determines whether this range is within or at the same text as a specified range.
        /// </summary>
        /// <param name="range"></param>

        public bool InRange(TextRange range)
        {
            return _range.InRange(range._range) == TomBoolean.True;
        }

        TomBoolean ITextRange.InRange(ITextRange range) => _range.InRange(range);

        /// <summary>
        ///  	Determines whether this range's story is the same as a specified range's story.
        /// </summary>
        /// <param name="range"></param>

        public bool InStory(TextRange range)
        {
            return _range.InStory(range._range) == TomBoolean.True;
        }

        TomBoolean ITextRange.InStory(ITextRange range) => _range.InStory(range);

        TomBoolean ITextRange.IsEqual(ITextRange range) => _range.IsEqual(range);

        /// <summary>
        /// Sets the start and end positions, and story values of the active selection, to those of this range.
        /// </summary>
        public void Select()
        {
            _range.Select();
        }

        /// <summary>
        /// Pastes text from a specified data object.
        /// </summary>
        /// <param name="expand">true to move the start of the range to the end of the pasted range; false to keep the start position at the start.</param>
        /// <param name="data">The IDataObject to paste. If null the clipboard is used.</param>
        /// <param name="format">The clipboard format to use in the paste operation. Zero is best format, which usually is RTF, but CF_UNICODETEXT and other formats are also possible. The default value is zero. For more information, see Clipboard Formats.</param>
        public void Paste(bool expand = false, IDataObject data = null, int format = 0)
        {
            var startPos = 0;
            if (expand)
                startPos = _range.Start;

            _range.Paste(data, format);

            if (expand)
                _range.Start = startPos;
        }

        void ITextRange.Paste(IDataObject data, int format) => _range.Paste(data, format);

        /// <summary>
        /// Determines if a data object can be pasted, using a specified format, into the current range.
        /// </summary>
        /// <param name="data">The IDataObject to paste. If null the clipboard is used.</param>
        /// <param name="format">The clipboard format to use in the paste operation. Zero is best format, which usually is RTF, but CF_UNICODETEXT and other formats are also possible. The default value is zero. For more information, see Clipboard Formats.</param>
        public bool CanPaste(IDataObject data = null, int format = 0)
            => _range.CanPaste(data, format) == TomBoolean.True;

        TomBoolean ITextRange.CanPaste(IDataObject data, int format) => _range.CanPaste(data, format);

        /// <summary>
        /// Determines whether the specified range can be edited.
        /// </summary>

        public bool CanEdit()
        {
            return _range.CanEdit();
        }

        /// <summary>
        /// Retrieves a pointer to the embedded object at the start of the specified range, that is, at cpFirst. The range must either be an insertion point or it must select only the embedded object.
        /// </summary>

        public object GetEmbeddedObject()
        {
            return _range.GetEmbeddedObject();
        }

        /// <summary>
        /// Scrolls the specified range into view.
        /// </summary>
        /// <param name="value">The start or end of the range to scroll into view.</param>
        public bool ScrollIntoView(RangePosition value)
        {
            return _range.ScrollIntoView(value) == 0;
        }

        int ITextRange.ScrollIntoView(RangePosition value) => _range.ScrollIntoView(value);

        /// <summary>
        /// Changes the range based on a specified point at or up through (depending on Extend) the point (x, y) aligned according to Type.
        /// </summary>
        /// <param name="x">Horizontal coordinate of the specified point, in absolute screen coordinates.</param>
        /// <param name="y">Vertical coordinate of the specified point, in absolute screen coordinates.</param>
        /// <param name="type">The end to move to the specified point.</param>
        /// <param name="extend">How to set the endpoints of the range. If Extend is zero (the default), the range is an insertion point at the specified point (or at the nearest point with selectable text). If Extend is 1, the end specified by Type is moved to the point and the other end is left alone.</param>
        public void SetPoint(int x, int y, RangePosition type, ExtendRange extend)
        {
            _range.SetPoint(x, y, type, extend);
        }

        /// <summary>
        ///  	Retrieves screen coordinates for the start or end character position in the text range, along with the inter-line position.
        /// </summary>
        /// <param name="type">Flag that indicates the position to retrieve. This parameter can include one value from each of the following tables. The default value is tomStart + TA_BASELINE + TA_LEFT.</param>
        /// <param name="x">The x-coordinate.</param>
        /// <param name="y">The y-coordinate.</param>
        public void GetPoint(GetPointType type, out int x, out int y)
        {
            _range.GetPoint(type, out x, out y);
        }

        /// <summary>
        ///     Changes the case of letters in this range according to the Type parameter.
        /// </summary>
        /// <param name="type"></param>
        public void ChangeCase(RangeChangeCase type)
        {
            _range.ChangeCase(type);
        }


        /// <summary>
        ///     Copies the TextRange to new DataObject.
        /// </summary>
        /// <returns>An Com IDataObject interface.</returns>
        public IDataObject CopyToDataObject()
        {
            var tdo = new TextRangeDataObject();
            _range.Copy(tdo);
            return tdo.DataObject;
        }

        /// <summary>
        ///     Returns the rtf contained within the TextRange
        /// </summary>
        /// <returns>A string containing rtf tags in the TextRange.</returns>
        public string ToRtf()
        {
            if (_range.Start == _range.End)
                return string.Empty;

            // TextRangeVariant is special helper class that's needed because the .Net interop does not support
            // the type of ByRef variant object that TextRange.Copy requires
            var trdo = new TextRangeDataObject();

            // TextRange.Copy passes a copy of it's stored rtf string and returns it in an IDataobject which we
            _range.Copy(trdo);
            if (trdo.DataObject != null)
            {
                //Return the rtf
                var rtfExtractor = new RichTextConverter();
                return rtfExtractor.ToString(trdo.DataObject);
            }

            return string.Empty;
        }

        /// <summary>
        ///     Returns the rtf contained within the TextRange as stream
        /// </summary>
        /// <returns>A stream containing rtf tags in the TextRange.</returns>
        public Stream ToStream()
        {
            if (_range.Start == _range.End)
                return new MemoryStream();

            // TextRangeVariant is special helper class that's needed because the .Net interop does not support
            // the type of ByRef variant object that TextRange.Copy requires
            var trdo = new TextRangeDataObject();

            // TextRange.Copy passes a copy of it's stored rtf string and returns it in an IDataobject which we
            _range.Copy(trdo);

            if (trdo.DataObject != null)
            {
                //Return the rtf
                var rtfExtractor = new RichTextConverter();
                var ms = new MemoryStream();
                try
                {
                    rtfExtractor.ToStream(trdo.DataObject, ms);
                    return ms;
                }
                catch
                {
                    ms.Dispose();
                    throw;
                }

            }

            return null;
        }

        /// <summary>
        ///     Replaces the text range with the supplied rtf.
        /// </summary>
        /// <param name="rtf">A string containing rtf tags to insert.</param>
        public void FromRtf(string rtf)
        {
            var ido = new RichTextDataObject(rtf);
            var start = _range.Start;

            _range.Paste(ido);
            _range.SetRange(_range.End, start);
        }

        /// <summary>
        ///     Replaces the text range with the supplied rtf.
        /// </summary>
        /// <param name="stream">A string containing rtf tags to insert.</param>
        public void FromStream(Stream stream)
        {
            var ido = new RichTextDataObject(stream);
            var start = _range.Start;

            _range.Paste(ido);
            _range.SetRange(_range.End, start);
        }

        /// <summary>
        ///     Copies the text to the clipboard.
        /// </summary>
        public void Copy() => _range.Copy(null);

        /// <summary>
        ///     Copies the text to a data object. If data is null the clipboard is used..
        /// </summary>
        /// <param name="data">An instance implementing the com IDataObject interface</param>
        void ITextRange.Copy(TextRangeDataObject data) => _range.Copy(data);

        /// <summary>
        ///     Cuts the plain or rich text to a data object or to the Clipboard. If data is null the clipboard is used..
        /// </summary>
        /// <param name="data">An instance implementing the com IDataObject interface</param>
        public void Cut(IDataObject data = null)
        {
            var tdo = new TextRangeDataObject { DataObject = data };
            _range.Cut(tdo);
        }

        void ITextRange.Cut(TextRangeDataObject data) => _range.Cut(data);

        /// <summary>
        ///     Mimics the DELETE and BACKSPACE keys, with and without the CTRL key depressed.
        /// </summary>
        /// <param name="unit"></param>
        /// <param name="count"></param>

        public int Delete(RangeUnit unit, int count = 1)
        {
            return _range.Delete(unit, count);
        }

        /// <summary>
        ///     Searches up to Count characters for the string, bstr, starting from the range's End cp.
        /// </summary>
        /// <param name="stringToFind"></param>
        /// <param name="count"></param>
        /// <param name="flags"></param>

        public int FindTextEnd(string stringToFind, int count, RangeFindTextFlags flags)
        {
            return _range.FindTextEnd(stringToFind, count, flags);
        }

        /// <summary>
        ///     Searches up to Count characters for the string, bstr, starting at the range's Start cp (cpFirst).
        /// </summary>
        /// <param name="stringToFind"></param>
        /// <param name="count"></param>
        /// <param name="flags"></param>

        public int FindTextStart(string stringToFind, int count = (int)RangeMoveDirection.Forward, RangeFindTextFlags flags = RangeFindTextFlags.MatchDefault)
        {
            return _range.FindTextStart(stringToFind, count, flags);
        }

        /// <summary>
        ///     Searches up to Count characters for the text given by bstr. The starting position and direction are also specified
        ///     by Count, and the matching criteria are given by Flags.
        /// </summary>
        /// <param name="stringToFind"></param>
        /// <param name="count"></param>
        /// <param name="flags"></param>

        public int FindText(string stringToFind, int count = (int)RangeMoveDirection.Forward, RangeFindTextFlags flags = RangeFindTextFlags.MatchDefault)
        {
            return _range.FindText(stringToFind, count, flags);
        }

        /// <summary>
        ///     Moves the range's end to the character position of the first character found that is in the set of characters
        ///     specified by charSet, provided that the character is found within Count characters of the range's end.
        /// </summary>
        /// <param name="charSet"></param>
        /// <param name="count"></param>

        public int MoveEndUntil(string charSet, RangeMoveDirection count = RangeMoveDirection.Forward)
        {
            return _range.MoveEndUntil(charSet, count);
        }

        int ITextRange.MoveEndUntil(ref object charSet, RangeMoveDirection count) => _range.MoveEndUntil(charSet, count);

        /// <summary>
        ///     Moves the start position of the range the position of the first character found that is in the set of characters
        ///     specified by charSet, provided that the character is found within Count characters of the start position.
        /// </summary>
        /// <param name="charSet"></param>
        /// <param name="count"></param>

        public int MoveStartUntil(string charSet, RangeMoveDirection count = RangeMoveDirection.Forward)
        {
            return _range.MoveStartUntil(charSet, count);
        }

        int ITextRange.MoveStartUntil(ref object charSet, RangeMoveDirection count) => _range.MoveStartUntil(charSet, count);

        /// <summary>
        ///     Searches up to Count characters for the first character in the set of characters specified by charSet. If a character
        ///     is found, the range is collapsed to that point. The start of the search and the direction are also specified by
        ///     Count.
        /// </summary>
        /// <param name="charSet"></param>
        /// <param name="count"></param>

        public int MoveUntil(string charSet, RangeMoveDirection count = RangeMoveDirection.Forward)
        {
            return _range.MoveUntil(charSet, count);
        }

        int ITextRange.MoveUntil(ref object charSet, RangeMoveDirection count) => _range.MoveUntil(charSet, count);

        /// <summary>
        ///     Moves the end of the range either Count characters or just past all contiguous characters that are found in the set
        ///     of characters specified by charSet, whichever is less.
        /// </summary>
        /// <param name="charSet"></param>
        /// <param name="count"></param>

        public int MoveEndWhile(string charSet, RangeMoveDirection count = RangeMoveDirection.Forward)
        {
            return _range.MoveEndWhile(charSet, count);
        }

        int ITextRange.MoveEndWhile(ref object charSet, RangeMoveDirection count) => _range.MoveEndWhile(charSet, count);

        /// <summary>
        ///     Moves the start position of the range either Count characters, or just past all contiguous characters that are
        ///     found in the set of characters specified by charSet, whichever is less.
        /// </summary>
        /// <param name="charSet"></param>
        /// <param name="count"></param>

        public int MoveStartWhile(string charSet, RangeMoveDirection count = RangeMoveDirection.Forward)
        {
            return _range.MoveStartWhile(charSet, count);
        }

        int ITextRange.MoveStartWhile(ref object charSet, RangeMoveDirection count) => _range.MoveStartWhile(charSet, count);

        /// <summary>
        ///     Starts at a specified end of a range and searches while the characters belong to the set specified by charSet and
        ///     while the number of characters is less than or equal to Count.
        /// </summary>
        /// <param name="charSet"></param>
        /// <param name="count"></param>

        public int MoveWhile(string charSet, RangeMoveDirection count = RangeMoveDirection.Forward)
        {
            return _range.MoveWhile(charSet, count);
        }

        int ITextRange.MoveWhile(ref object charSet, RangeMoveDirection count) => _range.MoveWhile(charSet, count);

        /// <summary>
        ///     Moves the end position of the range.
        /// </summary>
        /// <param name="unit"></param>
        /// <param name="count"></param>

        public int MoveEnd(RangeUnit unit = RangeUnit.Character, int count = 1)
        {
            return _range.MoveEnd(unit, count);
        }

        /// <summary>
        ///     Moves the start position of the range the specified number of units in the specified direction.
        /// </summary>
        /// <param name="unit"></param>
        /// <param name="count"></param>

        public int MoveStart(RangeUnit unit = RangeUnit.Character, int count = 1)
        {
            return _range.MoveStart(unit, count);
        }

        /// <summary>
        ///     Moves the insertion point forward or backward a specified number of units. If the range is non-degenerate, the range
        ///     is collapsed to an insertion point at either end, depending on Count, and then is moved.
        /// </summary>
        /// <param name="unit"></param>
        /// <param name="count"></param>

        public int Move(RangeUnit unit = RangeUnit.Character, int count = 1)
        {
            return _range.Move(unit, count);
        }

        /// <summary>
        ///     Moves this range's ends to the end of the last overlapping Unit in the range.
        /// </summary>
        /// <param name="unit"></param>
        /// <param name="extend"></param>

        public int EndOf(int unit, ExtendRange extend)
        {
            return _range.EndOf(unit, extend);
        }

        /// <summary>
        ///     Moves the range ends to the start of the first overlapping Unit in the range.
        /// </summary>
        /// <param name="unit"></param>
        /// <param name="extend"></param>

        /// <remarks>Search ITextRange.StartOf in a browser for more info.</remarks>
        public int StartOf(RangeUnit unit, ExtendRange extend)
        {
            return _range.StartOf(unit, extend);
        }

        /// <summary>
        ///     Collapses of expands the range by the specified unit and the supplied index;
        /// </summary>
        public void SetIndex(RangeUnit unit, int index, ExtendRange extend)
        {
            _range.SetIndex(unit, index, extend);
        }



        /// <summary>
        ///     Retrieves the story index of the Unit parameter at the specified range Start. The first unit in a story
        ///     has an index value of 1. The index of a unit is the same for all cps from that immediately preceding
        ///     the unit up to the last character in the Unit.
        /// </summary>
        /// <param name="unit">
        ///     One of the RangeUnit values that is indexed. For a list of possible unit values, see the discussion
        ///     under TextRange.
        /// </param>
        /// <returns>Receives the index value for the corresponding unit passed.</returns>
        /// <remarks>
        ///     <para>
        ///         The TextRange.GetIndex method retrieves the story index of a word,
        ///         line, sentence, paragraph, and so forth, at the range Start. unit specifies which kind of entity to index,
        ///         such as words (RangeUnit.Word), lines (RangeUnit.Line), sentences (RangeUnit.Sentence), or
        ///         paragraphs (RangeUnit.Paragraph). For example, TextRange.GetIndex returns a value equal to the line
        ///         number of the first line in the range. Note that for a range at the end of the story, TextRange.GetIndex,
        ///         returns the number of units in the story. Thus, you can get the number of words, lines, objects,
        ///         and so forth, in a story.
        ///     </para>
        ///     <para>
        ///         Note, the index value returned by the TextRange.GetIndex method is not valid if the
        ///         text is subsequently edited. Thus, users should be careful about using methods that return
        ///         index values, especially if the values are to be stored for any duration.
        ///         This is in contrast to a pointer to a range, which does remain valid when the text is edited.
        ///     </para>
        /// </remarks>
        public int GetIndex(RangeUnit unit)
        {
            return _range.GetIndex(unit);
        }

        /// <summary>
        ///     Expands this range so that any partial units it contains are completely contained.
        /// </summary>
        /// <param name="unit">
        ///     One of the RangeUnit values to include, if it is partially within the range.
        ///     The default value is RangeUnit.Word. For a list of the other RangeUnit values, see the discussion under TextRange.
        /// </param>
        /// <returns>The count of characters added to the range.</returns>
        /// <remarks>
        ///     For example, if an insertion point is at the beginning, the end, or within a word,
        ///     TextRange.Expand expands the range to include that word.
        ///     If the range already includes one word and part of another, TextRange.Expand expands
        ///     the range to include both words. TextRange.Expand expands the range to include the visible
        ///     portion of the range's story.
        /// </remarks>
        public int Expand(RangeUnit unit)
        {
            return _range.Expand(unit);
        }

        /// <summary>
        ///     Collapses the specified text range into a degenerate point at either the beginning or end of the range.
        /// </summary>
        /// <param name="start">One of the RangePosition flags specifying the end to collapse at.</param>

        public void Collapse(RangePosition start)
        {
            _range.Collapse(start);
        }


#pragma warning disable 659
        /// <summary>
        ///     Returns if two TextRange point to the same location
        /// </summary>
        public override bool Equals(object obj) => Equals(obj as TextRange);
#pragma warning restore 659

        /// <summary>
        ///     Determines whether this range has the same character positions and story as those of a specified range.
        /// </summary>
        /// <param name="other">The TextRange range that is compared to the current range.</param>


        public bool Equals(TextRange other)
        {
            if (other == null)
                return false;
            return _range.IsEqual(other._range) == TomBoolean.True;
        }

        /// <summary>
        ///     Disposes of resources the TextRange uses.
        /// </summary>
        public virtual void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        ///     Disposes of resources the TextRange uses.
        /// </summary>
        protected virtual void Dispose(bool disposing)
        {
            if (_range != null)
            {
                Marshal.ReleaseComObject(_range);
                _range = null;
            }
        }

        /// <summary>
        ///     Finalizer for the TextRange
        /// </summary>
        ~TextRange()
        {
#if (TRACK_DISPOSED)
            LevitJames.Core.Diagnostics.DisposeTracker.Add(_disposedSource);
#endif
            Dispose(disposing: false);
        }
    }
}