// © Copyright 2018 Levit & James, Inc.

using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

#pragma warning disable CS1591 // Missing XML comment for publicly visible type or member

namespace LevitJames.TextServices
{
    /// <summary>
    ///     Defines the set of rtf versions that cn be associated with a RichTextBox.
    /// </summary>

    public enum RichTextBoxVersion
    {
        /// <summary>
        ///     Finds the Latest version of the RichEdit Dll library to use
        /// </summary>

        Best = 0,

        /// <summary>
        ///     Office 2000, Windows ME/2000/XP/Vista
        /// </summary>
        /// <remarks>Version is 3.1 on Vista</remarks>
        Version30 = 3,

        /// <summary>
        ///     Office XP
        /// </summary>

        Version40,

        /// <summary>
        ///     Windows XP SP1, Tablet and Vista / Office 2003
        /// </summary>
        /// <remarks>4.1 for Windows XP SP1, Tablet and Vista</remarks>
        Version50,

        /// <summary>
        ///     Office 2007, Encarta Math Calculator
        /// </summary>

        Version60,

        /// <summary>
        ///     Office 2010
        /// </summary>

        Version70
    }


    /// <summary>
    ///     Defines the set of values returned from the
    /// </summary>

    public enum OutlineLevel
    {
        Unknown = 0,
        Level1 = 1,
        Level2 = 2,
        Level3 = 3,
        Level4 = 4,
        Level5 = 5,
        Level6 = 6,
        Level7 = 7,
        Level8 = 8,
        Level9 = 9,
        BodyText = 10
    }


    /// <summary>
    ///     Defines the set of boolean values used by the TextRange, TextSelection and TextParagraph objects.
    /// </summary>

    [CompilerGenerated]
    public enum TomBoolean
    {
        Toggle = -9999998,
        False = 0,
        True = -1,
        Undefined = -9999999
    }


    /// <summary>
    ///     Defines the set of character match sets used by the TextRange, TextSelection objects for assisting moving through
    ///     the range.
    /// </summary>

    [Flags]
    [CompilerGenerated]
    public enum CharacterMatchSets
    {
        AlphaCharacters = 0x100,
        BlankCharacters = 0x40,
        ControlCharacters = 0x20,
        DecimalDigits = 4,
        DefinedCharacter = 0x400,
        HexadecimalDigits = 0x80,
        Lowercase = 2,
        Punctuation = 0x10,
        SpaceCharacters = 8,
        Uppercase = 1
    }


    /// <summary>
    ///     Defines the set of custom forecolor values that can be returned or set from the TextFont.ForeColor property.
    /// </summary>
    /// <remarks>
    ///     These values are additional to the standard set of color values you can apply to the TextFont.ForeColor
    ///     property.
    /// </remarks>
    [CompilerGenerated]
    public enum TextColor
    {
        AutoColor = -9999997,
        ColorUndefined = -9999999
    }


    /// <summary>
    ///     Defines the set options for extending or moving the end or start point of a TextRange object.
    /// </summary>
    /// <remarks>
    ///     Flag that indicates how to change the selection. If Extend =  Move, the method collapses
    ///     the selection to an insertion point. If Extend = Extend, the method moves the active end and leaves
    ///     the other end alone. The default value is zero.
    /// </remarks>
    [CompilerGenerated]
    public enum ExtendRange
    {
        /// <summary>
        ///     Collapses the range to an insertion point
        /// </summary>

        Move = 0,

        /// <summary>
        ///     Moves the active end and leaves
        ///     the other end alone. The default value is zero.
        /// </summary>

        Extend = 1
    }


    /// <summary>
    ///     Defines the set of flags to use when opening or closing a TextDocument.
    /// </summary>

    [CompilerGenerated]
    [Flags]
    public enum FileSaveOpenFlags
    {
        None = 0,
        CreateAlways = 0x20,
        CreateNew = 0x10,

        [EditorBrowsable(EditorBrowsableState.Never)]
        Html = 3,
        OpenAlways = 0x40,
        OpenExisting = 0x30,
        PasteFile = 0x1000,
        ReadOnly = 0x100,
        Rtf = 1,
        ShareDenyRead = 0x200,
        ShareDenyWrite = 0x400,
        Text = 2,
        TruncateExisting = 80,

        [EditorBrowsable(EditorBrowsableState.Never)]
        WordDocument = 4
    }


    /// <summary>
    ///     Defines the set of values to set for the TextFont.Animation property.
    /// </summary>

    [CompilerGenerated]
    public enum FontAnimation
    {
        BlinkingBackground = 2,
        LasVegasLights = 1,
        MarchingBlackAnts = 4,
        MarchingRedAnts = 5,
        NoAnimation = 0,
        Shimmer = 6,
        SparkleText = 3,
        WipeDown = 7,
        WipeRight = 8
    }


    /// <summary>
    ///     Defines the set of values to set for the TextFont.Underline property.
    /// </summary>

    [CompilerGenerated]
    public enum FontUnderline
    {
        None = 0,
        Single = 1,
        Words = 2,
        Double = 3,
        Dotted = 4,
        Undefined = -9999999
    }


    /// <summary>
    ///     Defines the set of values to set for the TextFont.Weight property.
    /// </summary>

    [CompilerGenerated]
    public enum FontWeight
    {
        DontCare = 0,
        Thin = 100,
        ExtraLight = 200,
        Light = 300,
        Normal = 400,
        Medium = 500,
        SemiBold = 600,
        Bold = 700,
        ExtraBold = 800,
        Heavy = 900
    }


    /// <summary>
    ///     Defines the set of flags used to define which location values to return from a TextRange.GetPoint property.
    /// </summary>

    [Flags]
    [CompilerGenerated]
    public enum GetPointType
    {
        End = 0,
        Left = 0,
        Top = 0,
        Right = 2,
        Bottom = 8,
        Center = 6,
        Start = 0x20,
        Baseline = 0x18
    }


    /// <summary>
    ///     Defines the set of tab leader values to set for setting the various TextParagraph tab Properties.Resources.
    /// </summary>

    [CompilerGenerated]
    public enum ParagraphAddTabLeader
    {
        Spaces = 0,
        Dots = 1,
        Dashes = 2,
        Lines = 3
    }


    /// <summary>
    ///     Defines the set of paragraph alignment values used in the TextParagraph.Alignment property.
    /// </summary>

    [CompilerGenerated]
    public enum ParagraphAlignment
    {
        /// <summary>
        ///     Text aligns with the left margin.
        /// </summary>

        Left = 0,

        /// <summary>
        ///     Text is centered between the margins.
        /// </summary>

        Center = 1,

        /// <summary>
        ///     Text aligns with the right margin.
        /// </summary>

        Right = 2,

        /// <summary>
        ///     Text starts at the left margin and, if the line extends beyond the right margin, all the spaces in the line are
        ///     adjusted to be even.
        /// </summary>

        Justify = 3
    }


    /// <summary>
    ///     Defines the set of values for setting the line spacing rules of a TextParagraph object.
    /// </summary>

    [CompilerGenerated]
    public enum ParagraphLineSpacingRule
    {
        /// <summary>
        ///     The line-spacing value is ignored.
        /// </summary>

        Single = 0,

        /// <summary>
        ///     The line-spacing value is ignored.
        /// </summary>

        OnePointFive = 1,

        /// <summary>
        ///     The line-spacing value is ignored.
        /// </summary>

        Double = 2,

        /// <summary>
        ///     The line-spacing value specifies the spacing, in floating-point points, from one line to the next.
        ///     However, if the value is less than single spacing, the control displays single-spaced text.
        /// </summary>

        AtLeast = 3,

        /// <summary>
        ///     The line-spacing value specifies the exact spacing, in floating-point points, from one line to the next (even if
        ///     the value is less than single spacing).
        /// </summary>

        Exactly = 4,

        /// <summary>
        ///     The line-spacing value specifies the exact spacing, in floating-point points, from one line to the next (even if
        ///     the value is less than single spacing).
        /// </summary>

        Multiple = 5
    }


    /// <summary>
    ///     Defines the set of values for setting the alignment of a list in a TextParagraph object.
    /// </summary>

    [CompilerGenerated]
    public enum ParagraphListAlignment
    {
        Left = 0,
        Center = 1,
        Right = 2
    }


    /// <summary>
    ///     Defines the set of values that set the bullet style of a list in a TextParagraph object.
    /// </summary>

    [CompilerGenerated]
    public enum ParagraphListType
    {
        None = 0,
        Bullet = 1,
        NumberAsArabic = 2,
        NumberAsLowerCaseLetter = 3,
        NumberAsLowerCaseRoman = 5,
        NumberAsSequence = 7,
        NumberAsUpperCaseLetter = 4,
        NumberAsUpperCaseRoman = 6,
        Parentheses = 0x10000,
        Period = 0x20000,
        Plain = 0x30000
    }


    /// <summary>
    ///     Defines the set of values for setting the alignment of tabs in a TextParagraph object.
    /// </summary>

    [CompilerGenerated]
    public enum ParagraphTabAlignment
    {
        Left = 0,
        Center = 1,
        Right = 2,
        Decimal = 3,
        Justify = 3,
        Bar = 4
    }


    /// <summary>
    ///     Defines the set of additional values for retrieving a tab from the TextParagraph.GetTab method.
    /// </summary>

    [CompilerGenerated]
    public enum ParagraphTabIndex
    {
        Back = -3,
        Next = -2,
        Here = -1
    }


    /// <summary>
    ///     Defines the set of values for changing the case of text on a TextRange object.
    /// </summary>

    [CompilerGenerated]
    public enum RangeChangeCase
    {
        LowerCase = 0,
        UpperCase = 1,
        TitleCase = 2,
        SentenceCase = 4,
        ToggleCase = 5
    }


    /// <summary>
    ///     Defines the set of flags to set when finding text in a TextRage object.
    /// </summary>

    [CompilerGenerated]
    [Flags]
    public enum RangeFindTextFlags
    {
        MatchDefault = 0,
        MatchWord = 2,
        MatchCase = 4,
        MatchPattern = 8
    }


    /// <summary>
    ///     Defines the set of move direction values used when navigating though a TextRange object.
    /// </summary>

    [CompilerGenerated]
    public enum RangeMoveDirection
    {
        Backward = -1073741823,
        Forward = 0x3FFFFFFF
    }


    /// <summary>
    ///     Defines the set of range position values a TextRange object to act upon.
    /// </summary>

    [CompilerGenerated]
    public enum RangePosition
    {
        End = 0,
        Start = 0x20
    }


    /// <summary>
    ///     Defines the set of story type values for a TextRange object. (Not Implemented)
    /// </summary>

    [CompilerGenerated]
    public enum RangeStoryType
    {
        Unknown = 0,
        MainText = 1,
        Footnotes = 2,
        Endnotes = 3,
        Comments = 4,
        TextFrame = 5,
        EvenPagesHeader = 6,
        PrimaryHeader = 7,
        EvenPagesFooter = 8,
        PrimaryFooter = 9,
        FirstPageHeader = 10,
        FirstPageFooter = 11
    }


    /// <summary>
    ///     Defines the set of values for the type of TextRange unit to navigate with.
    /// </summary>

    [CompilerGenerated]
    public enum RangeUnit
    {
        Character = 1,
        Word = 2,
        Sentence = 3,
        Paragraph = 4,
        Line = 5,
        Story = 6,
        Screen = 7,
        Section = 8,
        Column = 9,
        Row = 10,
        Window = 11,
        Cell = 12,
        CharFormat = 13,
        ParaFormat = 14,
        Table = 15,
        Object = 0x10
    }


    /// <summary>
    ///     Defines the set of values for resetting the character properties of a TextFont object.
    /// </summary>

    [CompilerGenerated]
    public enum ResetTextFontValue
    {

        ApplyLater = 1,
        ApplyNow = 0,
        ApplyRtfDocProps = 0x4000,
        ApplyTemp = 4,

        /// <summary>
        ///     System uses the default values defined by the RTF/plain control word.
        /// </summary>

        Default = -9999996,

        /// <summary>
        ///     Sets all properties to undefined values.
        /// </summary>

        Undefined = -9999999,


    }


    /// <summary>
    ///     Defines the set of values for resetting the character properties of a TextParagraph object.
    /// </summary>

    [CompilerGenerated]
    public enum ResetParagraphValue
    {
        /// <summary>
        ///     Used for paragraph formatting that is defined by the RTF \pard (that is, the paragraph default) control word.
        /// </summary>

        Default = -9999996,

        /// <summary>
        ///     Used for all undefined values. The Undefined value is only valid for duplicate (clone) TextParagraph objects.
        /// </summary>

        Undefined = -9999999
    }


    /// <summary>
    ///     Defines the set of values for the calling the TextSelection HomeKey and EndKey members.
    /// </summary>

    [CompilerGenerated]
    public enum SelectionKey
    {
        /// <summary>
        ///     Depending on Extend, it moves either the insertion point or the active end to the end of the last line in the
        ///     selection. This is the default.
        /// </summary>

        Line = 5,

        /// <summary>
        ///     Depending on Extend, it moves either the insertion point or the active end to the end of the last line in the
        ///     story.
        /// </summary>

        Story = 6,

        /// <summary>
        ///     Depending on Extend, it moves either the insertion point or the active end to the end of the last column in the
        ///     selection. This is available only if the Text Object Model (TOM) engine supports tables.
        /// </summary>

        Column = 9,

        /// <summary>
        ///     Depending on Extend, it moves either the insertion point or the active end to the end of the last row in the
        ///     selection. This is available only if the TOM engine supports tables.
        /// </summary>

        Row = 10
    }


    /// <summary>
    ///     Defines the set of flags that describe state of the TextSelection object.
    /// </summary>

    [CompilerGenerated]
    [Flags]
    public enum SelectionFlags
    {
        StartActive = 0x1,
        AtEndOfLine = 0x2,
        Overtype = 0x4,
        Active = 0x8,
        Replace = 0x10
    }


    /// <summary>
    ///     Defines the set of values for navigating horizontaly through a TextSelection object.
    /// </summary>

    [CompilerGenerated]
    public enum SelectionMoveHorizontal
    {
        /// <summary>
        ///     Move one character position to the left or right. This is the default
        /// </summary>

        Character = 1,

        /// <summary>
        ///     Move one word to the left or right.
        /// </summary>

        Word = 2
    }


    /// <summary>
    ///     Defines the set of values for navigating vertically through a TextSelection object.
    /// </summary>

    [CompilerGenerated]
    public enum SelectionMoveVertical
    {
        Line = 5,
        Paragraph = 4,
        Screen = 7,
        Window = 11
    }


    /// <summary>
    ///     Defines the set of values that describe the type selection in a TextSelection object.
    /// </summary>

    [CompilerGenerated]
    public enum SelectionType
    {
        None = 0,
        InsertionPoint = 1,
        Normal = 2,
        Frame = 3,
        Column = 4,
        Row = 5,
        Block = 6,
        InlineShape = 7,
        Shape = 8
    }


    /// <summary>
    ///     Defines the set of built in styles for a TextFont object.
    /// </summary>
    /// <remarks>It appears that only some of these defined values are actually implemented by the RichEdit control.</remarks>
    [CompilerGenerated]
    public enum BuiltInStyles
    {
        Normal = -1,
        Heading1 = -2,
        Heading2 = -3,
        Heading3 = -4,
        Heading4 = -5,
        Heading5 = -6,
        Heading6 = -7,
        Heading7 = -8,
        Heading8 = -9,
        Heading9 = -10,
        Index1 = -11,
        Index2 = -12,
        Index3 = -13,
        Index4 = -14,
        Index5 = -15,
        Index6 = -16,
        Index7 = -17,
        Index8 = -18,
        Index9 = -19,
        TableOfContents1 = -20,
        TableOfContentsC2 = -21,
        TableOfContents3 = -22,
        TableOfContents4 = -23,
        TableOfContents5 = -24,
        TableOfContents6 = -25,
        TableOfContents7 = -26,
        TableOfContents8 = -27,
        TableOfContents9 = -28,
        NormalIndent = -29,
        FootnoteText = -30,
        AnnotationText = -31,
        Header = -32,
        Footer = -33,
        IndexHeading = -34,
        Caption = -35,
        TableofFigures = -36,
        EnvelopeAddress = -37,
        EnvelopeReturn = -38,
        FootnoteReference = -39,
        AnnotationReference = -40,
        LineNumber = -41,
        PageNumber = -42,
        EndnoteReference = -43,
        EndnoteText = -44,
        TableofAuthorities = -45,
        MacroText = -46,
        TableOfContentsHeading = -47,
        List = -48,
        SListBullet = -49,
        ListNumber = -50,
        List2 = -51,
        List3 = -52,
        List4 = -53,
        List5 = -54,
        ListBullet2 = -55,
        ListBullet3 = -56,
        ListBullet4 = -57,
        ListBullet5 = -58,
        ListNumber2 = -59,
        ListNumber3 = -60,
        ListNumber4 = -61,
        ListNumber5 = -62,
        Title = -63,
        Closing = -64,
        Signature = -65,
        BodyTextIndent = -68,
        ListContinue = -69,
        ListContinue2 = -70,
        ListContinue3 = -71,
        ListContinue4 = -72,
        ListContinue5 = -73,
        MessageHeader = -74,
        Subtitle = -75,
        Salutation = -76,
        Date = -77,
        BodyTextFirstIndent = -78,
        BodyTextFirstIndent2 = -79,
        NoteHeading = -80,
        BodyText2 = -81,
        BodyText3 = -82,
        BodyTextIndent2 = -83,
        BodyTextIndent3 = -84,
        BlockQuotation = -85,
        Hyperlink = -86,
        HyperlinkFollowed = -87
    }
}