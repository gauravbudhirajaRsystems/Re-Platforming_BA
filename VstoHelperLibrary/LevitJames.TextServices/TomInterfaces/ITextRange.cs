// © Copyright 2018 Levit & James, Inc.

using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace LevitJames.TextServices
{
    /// <summary>
    ///     Represents a Text Object Model interface ITextRange. See the MSDN ITextRange documentation for more details.
    /// </summary>
    /// <remarks>
    ///     This is a Com interface so Marshal.ReleaseComObject should be used to free the object when you have finished
    ///     with it.
    /// </remarks>
    [CompilerGenerated]
    [SuppressMessage("Microsoft.Naming", "CA1715:IdentifiersShouldHaveCorrectPrefix", MessageId = "I")]
    [ComImport]
    [Guid("8CC497C2-A1DF-11CE-8098-00AA0047BE5D")]
    [TypeLibType(flags: 0x10C0)]
    [DefaultMember("Text")]
    internal interface ITextRange
    {
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
        [DispId(dispId: 0)]
        string Text { get; set; }

        /// <summary>
        ///     Gets or sets the character at the range's start position.
        /// </summary>


        /// <remarks>
        ///     Similarly, setting TextRange.Char overwrites the first character with the character,
        ///     Char. Note that the characters retrieved and set by these methods are int32 variables, which hide the way that they
        ///     are stored in the
        ///     backing store (as bytes, words, variable-length, and so forth), and they do not require using a Unicode
        ///     System.String.
        ///     The Char property, which can do most things that a characters collection can, has two big advantages:
        ///     It can reference any character in the parent story instead of being limited to the parent range.
        ///     It is significantly faster, since System.Char are involved instead of range objects.
        ///     Accordingly, the Text Object Model (TOM) does not support a characters collection.
        /// </remarks>
        [DispId(dispId: 0x201)]
        int Char { get; set; }

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
        [DispId(dispId: 0x202)]
        ITextRange Duplicate { get; }

        /// <summary>
        /// </summary>



        [DispId(dispId: 0x203)]
        ITextRange FormattedText { get; set; }

        /// <summary>
        ///     Gets or sets the specified range Start.
        /// </summary>

        /// <returns>The start position of the range object.</returns>

        [DispId(dispId: 0x204)]
        int Start { get; set; }

        /// <summary>
        ///     Gets or sets the specified range End.
        /// </summary>

        /// <returns>The end position of the range object.</returns>

        [DispId(dispId: 0x205)]
        int End { get; set; }


        /// <summary>
        ///     Returns the TextFont object with the character attributes of the specified range.
        /// </summary>

        /// <returns>A TextFont instance.</returns>
        /// <remarks>
        ///     For plain-text controls, these objects do not vary from range to range,
        ///     but in rich-text solutions, they do. See the section on TextFont for further details.
        /// </remarks>
        [DispId(dispId: 0x206)]
        ITextFont Font { get; set; }

        /// <summary>
        ///     Returns an TextParagraph object with the paragraph attributes of the specified range
        /// </summary>

        /// <returns>A TextParagraph instance.</returns>

        [DispId(dispId: 0x207)]
        ITextPara Para { get; set; }

        /// <summary>
        ///     Returns the count of characters in the specified range's story.
        /// </summary>

        /// <returns>A count of the characters in the specified range's story</returns>

        [DispId(dispId: 520)]
        int StoryLength { get; }

        /// <summary>
        ///     Returns the type of the specified range's story.
        /// </summary>

        /// <returns>One of the values defined in RangeStoryType.</returns>
        /// <remarks>Currently the RichEdit control only supports RangeStoryType.MainText</remarks>
        [DispId(dispId: 0x209)]
        RangeStoryType StoryType { get; }

        /// <summary>
        ///     Collapses the specified text range into a degenerate point at either the beginning or end of the range.
        /// </summary>
        /// <param name="start">One of the RangePosition flags specifying the end to collapse at.</param>

        [DispId(dispId: 0x210)]
        void Collapse([In] RangePosition start);

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
        [DispId(dispId: 0x211)]
        int Expand([In] RangeUnit unit = RangeUnit.Word);

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
        [DispId(dispId: 530)]
        int GetIndex([In] RangeUnit unit);


        [DispId(dispId: 0x213)]
        void SetIndex([In] RangeUnit unit, [In] int index, [In] ExtendRange extend);


        [DispId(dispId: 0x214)]
        void SetRange([In] int active, [In] int other);


        [return: MarshalAs(UnmanagedType.I4)]
        [DispId(dispId: 0x215)]
        TomBoolean InRange([In] [MarshalAs(UnmanagedType.Interface)]
                           ITextRange range);


        [return: MarshalAs(UnmanagedType.I4)]
        [DispId(dispId: 0x216)]
        TomBoolean InStory([In] [MarshalAs(UnmanagedType.Interface)]
                           ITextRange range);


        [return: MarshalAs(UnmanagedType.I4)]
        [DispId(dispId: 0x217)]
        TomBoolean IsEqual([In] [MarshalAs(UnmanagedType.Interface)]
                           ITextRange range);


        [DispId(dispId: 0x218)]
        void Select();


        [DispId(dispId: 0x219)]
        int StartOf([In] RangeUnit unit, [In] ExtendRange extend);


        [DispId(dispId: 0x220)]
        int EndOf([In] int unit, [In] ExtendRange extend);


        [DispId(dispId: 0x221)]
        int Move([In] RangeUnit unit, [In] int count = 1);


        [DispId(dispId: 0x222)]
        int MoveStart([In] RangeUnit unit = RangeUnit.Character, [In] int count = 1);


        [DispId(dispId: 0x223)]
        int MoveEnd([In] RangeUnit unit = RangeUnit.Character, [In] int count = 1);


        [DispId(dispId: 0x224)]
        int MoveWhile([In][MarshalAs(UnmanagedType.Struct)] ref object charSet,
                      [In] RangeMoveDirection count = RangeMoveDirection.Forward);


        [DispId(dispId: 0x225)]
        int MoveStartWhile([In][MarshalAs(UnmanagedType.Struct)] ref object charSet,
                           [In] RangeMoveDirection count = RangeMoveDirection.Forward);


        [DispId(dispId: 550)]
        int MoveEndWhile([In][MarshalAs(UnmanagedType.Struct)] ref object charSet,
                         [In] RangeMoveDirection count = RangeMoveDirection.Forward);


        [DispId(dispId: 0x227)]
        int MoveUntil([In][MarshalAs(UnmanagedType.Struct)] ref object charSet,
                      [In] RangeMoveDirection count = RangeMoveDirection.Forward);


        [DispId(dispId: 0x228)]
        int MoveStartUntil([In][MarshalAs(UnmanagedType.Struct)] ref object charSet,
                           [In] RangeMoveDirection count = RangeMoveDirection.Forward);


        [DispId(dispId: 0x229)]
        int MoveEndUntil([In][MarshalAs(UnmanagedType.Struct)] ref object charSet,
                         [In] RangeMoveDirection count = RangeMoveDirection.Forward);


        [DispId(dispId: 560)]
        int FindText([In][MarshalAs(UnmanagedType.BStr)] string stringToFind,
                     [In] int count = (int)RangeMoveDirection.Forward,
                     [In] RangeFindTextFlags flags = RangeFindTextFlags.MatchDefault);


        [DispId(dispId: 0x231)]
        int FindTextStart([In][MarshalAs(UnmanagedType.BStr)] string stringToFind,
                          [In] int count = (int)RangeMoveDirection.Forward,
                          [In] RangeFindTextFlags flags = RangeFindTextFlags.MatchDefault);


        [DispId(dispId: 0x232)]
        int FindTextEnd([In][MarshalAs(UnmanagedType.BStr)] string stringToFind,
                        [In] int count = (int)RangeMoveDirection.Forward,
                        [In] RangeFindTextFlags flags = RangeFindTextFlags.MatchDefault);


        [DispId(dispId: 0x233)]
        int Delete([In] RangeUnit unit, [In] int count = 1);


        [DispId(dispId: 0x234)]
        void Cut([In] [Out] [MarshalAs(UnmanagedType.CustomMarshaler, MarshalTypeRef = typeof(TextRangeVariantMarshaler))]
                 TextRangeDataObject data = null);


        [DispId(dispId: 565)]
        void Copy([In] [Out] [MarshalAs(UnmanagedType.CustomMarshaler, MarshalTypeRef = typeof(TextRangeVariantMarshaler))]
                  TextRangeDataObject data = null);

        [DispId(dispId: 566)]
        void Paste(
            [In] [Out] [MarshalAs(UnmanagedType.CustomMarshaler, MarshalTypeRef = typeof(TextRangeVariantMarshaler))]
            IDataObject data = null, [In] int format = 0);


        [DispId(dispId: 0x237)]
        TomBoolean CanPaste([In] [Out] [MarshalAs(UnmanagedType.CustomMarshaler, MarshalTypeRef = typeof(TextRangeVariantMarshaler))]
                            IDataObject data = null, [In] int format = 0);


        [return: MarshalAs(UnmanagedType.I4)]
        [DispId(dispId: 0x238)]
        bool CanEdit();


        [DispId(dispId: 0x239)]
        void ChangeCase([In] RangeChangeCase type);


        [DispId(dispId: 0x240)]
        void GetPoint([In] GetPointType type, out int x, out int y);


        [DispId(dispId: 0x241)]
        void SetPoint([In] int x, [In] int y, [In] RangePosition type, [In] ExtendRange extend);


        //Return HRESULT not void
        //[DispId(dispId: 0x242)]
        //[PreserveSig]
        //void ScrollIntoView([In] RangePosition value);
        [DispId(dispId: 0x242)]
        [PreserveSig]
        int ScrollIntoView([In] RangePosition value);


        [return: MarshalAs(UnmanagedType.IUnknown)]
        [DispId(dispId: 0x243)]
        object GetEmbeddedObject();
    }
}