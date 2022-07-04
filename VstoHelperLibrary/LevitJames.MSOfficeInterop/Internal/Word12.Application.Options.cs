// © Copyright 2018 Levit & James, Inc.

using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace LevitJames.MSOffice.Internal
{
    //7 to get past the IDispatch members
    //The ShowWindowsInTaskbar member is member number 201 so that's 201-7 gives us 194 hence _VtblGap7_194 
    //Note: REMEMBER that read/write properties take 2 vtable slots.

    [ComImport]
    [Guid("000209B7-0000-0000-C000-000000000046")]
    [TypeLibType(flags: 0x10C0)]
    internal interface Options12 //316-7
    {
        [MethodImpl(MethodCodeType = MethodCodeType.Runtime)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        void _VtblGap7_311();

        //	<DispId(&H3E8)> _
        //	ReadOnly Property Application As <MarshalAs(UnmanagedType.Interface)> Application
        //	<DispId(&H3E9)> _
        //	ReadOnly Property Creator As Integer
        //	<DispId(&H3EA)> _
        //	ReadOnly Property Parent As <MarshalAs(UnmanagedType.IDispatch)> Object
        //	<DispId(1)> _
        //	Property AllowAccentedUppercase As Boolean
        //	<DispId(&H11)> _
        //	Property WPHelp As Boolean
        //	<DispId(&H12)> _
        //	Property WPDocNavKeys As Boolean
        //	<DispId(&H13)> _
        //	Property Pagination As Boolean
        //	<DispId(20)> _
        //	Property BlueScreen As Boolean
        //	<DispId(&H15)> _
        //	Property EnableSound As Boolean
        //	<DispId(&H16)> _
        //	Property ConfirmConversions As Boolean
        //	<DispId(&H17)> _
        //	Property UpdateLinksAtOpen As Boolean
        //	<DispId(&H18)> _
        //	Property SendMailAttach As Boolean
        //	<DispId(&H1A)> _
        //	Property MeasurementUnit As WdMeasurementUnits
        //	<DispId(&H1B)> _
        //	Property ButtonFieldClicks As Integer
        //	<DispId(&H1C)> _
        //	Property ShortMenuNames As Boolean
        //	<DispId(&H1D)> _
        //	Property RTFInClipboard As Boolean
        //	<DispId(30)> _
        //	Property UpdateFieldsAtPrint As Boolean
        //	<DispId(&H1F)> _
        //	Property PrintProperties As Boolean
        //	<DispId(&H20)> _
        //	Property PrintFieldCodes As Boolean
        //	<DispId(&H21)> _
        //	Property PrintComments As Boolean
        //	<DispId(&H22)> _
        //	Property PrintHiddenText As Boolean
        //	<DispId(&H23)> _
        //	ReadOnly Property EnvelopeFeederInstalled As Boolean
        //	<DispId(&H24)> _
        //	Property UpdateLinksAtPrint As Boolean
        //	<DispId(&H25)> _
        //	Property PrintBackground As Boolean
        //	<DispId(&H26)> _
        //	Property PrintDrawingObjects As Boolean
        //	<DispId(&H27)> _
        //	Property DefaultTray As <MarshalAs(UnmanagedType.BStr)> String
        //	<DispId(40)> _
        //	Property DefaultTrayID As Integer
        //	<DispId(&H29)> _
        //	Property CreateBackup As Boolean
        //	<DispId(&H2A)> _
        //	Property AllowFastSave As Boolean
        //	<DispId(&H2B)> _
        //	Property SavePropertiesPrompt As Boolean
        //	<DispId(&H2C)> _
        //	Property SaveNormalPrompt As Boolean
        //	<DispId(&H2D)> _
        //	Property SaveInterval As Integer
        //	<DispId(&H2E)> _
        //	Property BackgroundSave As Boolean
        //	<DispId(&H39)> _
        //	Property InsertedTextMark As WdInsertedTextMark
        //	<DispId(&H3A)> _
        //	Property DeletedTextMark As WdDeletedTextMark
        //	<DispId(&H3B)> _
        //	Property RevisedLinesMark As WdRevisedLinesMark
        //	<DispId(60)> _
        //	Property InsertedTextColor As WdColorIndex
        //	<DispId(&H3D)> _
        //	Property DeletedTextColor As WdColorIndex
        //	<DispId(&H3E)> _
        //	Property RevisedLinesColor As WdColorIndex
        //	<DispId(&H41)> _
        //	Property DefaultFilePath(ByVal Path As WdDefaultFilePath) As <MarshalAs(UnmanagedType.BStr)> String
        //	<DispId(&H42)> _
        //	Property Overtype As Boolean
        //	<DispId(&H43)> _
        //	Property ReplaceSelection As Boolean
        //	<DispId(&H44)> _
        //	Property AllowDragAndDrop As Boolean
        //	<DispId(&H45)> _
        //	Property AutoWordSelection As Boolean
        //	<DispId(70)> _
        //	Property INSKeyForPaste As Boolean
        //	<DispId(&H47)> _
        //	Property SmartCutPaste As Boolean
        //	<DispId(&H48)> _
        //	Property TabIndentKey As Boolean
        //	<DispId(&H49)> _
        //	Property PictureEditor As <MarshalAs(UnmanagedType.BStr)> String
        //	<DispId(&H4A)> _
        //	Property AnimateScreenMovements As Boolean
        //	<DispId(&H4B)> _
        //	Property VirusProtection As Boolean
        //	<DispId(&H4C)> _
        //	Property RevisedPropertiesMark As WdRevisedPropertiesMark
        //	<DispId(&H4D)> _
        //	Property RevisedPropertiesColor As WdColorIndex
        //	<DispId(&H4F)> _
        //	Property SnapToGrid As Boolean
        //	<DispId(80)> _
        //	Property SnapToShapes As Boolean
        //	<DispId(&H51)> _
        //	Property GridDistanceHorizontal As Single
        //	<DispId(&H52)> _
        //	Property GridDistanceVertical As Single
        //	<DispId(&H53)> _
        //	Property GridOriginHorizontal As Single
        //	<DispId(&H54)> _
        //	Property GridOriginVertical As Single
        //	<DispId(&H56)> _
        //	Property InlineConversion As Boolean
        //	<DispId(&H57)> _
        //	Property IMEAutomaticControl As Boolean
        //	<DispId(250)> _
        //	Property AutoFormatApplyHeadings As Boolean
        //	<DispId(&HFB)> _
        //	Property AutoFormatApplyLists As Boolean
        //	<DispId(&HFC)> _
        //	Property AutoFormatApplyBulletedLists As Boolean
        //	<DispId(&HFD)> _
        //	Property AutoFormatApplyOtherParas As Boolean
        //	<DispId(&HFE)> _
        //	Property AutoFormatReplaceQuotes As Boolean
        //	<DispId(&HFF)> _
        //	Property AutoFormatReplaceSymbols As Boolean
        //	<DispId(&H100)> _
        //	Property AutoFormatReplaceOrdinals As Boolean
        //	<DispId(&H101)> _
        //	Property AutoFormatReplaceFractions As Boolean
        //	<DispId(&H102)> _
        //	Property AutoFormatReplacePlainTextEmphasis As Boolean
        //	<DispId(&H103)> _
        //	Property AutoFormatPreserveStyles As Boolean
        //	<DispId(260)> _
        //	Property AutoFormatAsYouTypeApplyHeadings As Boolean
        //	<DispId(&H105)> _
        //	Property AutoFormatAsYouTypeApplyBorders As Boolean
        //	<DispId(&H106)> _
        //	Property AutoFormatAsYouTypeApplyBulletedLists As Boolean
        //	<DispId(&H107)> _
        //	Property AutoFormatAsYouTypeApplyNumberedLists As Boolean
        //	<DispId(&H108)> _
        //	Property AutoFormatAsYouTypeReplaceQuotes As Boolean
        //	<DispId(&H109)> _
        //	Property AutoFormatAsYouTypeReplaceSymbols As Boolean
        //	<DispId(&H10A)> _
        //	Property AutoFormatAsYouTypeReplaceOrdinals As Boolean
        //	<DispId(&H10B)> _
        //	Property AutoFormatAsYouTypeReplaceFractions As Boolean
        //	<DispId(&H10C)> _
        //	Property AutoFormatAsYouTypeReplacePlainTextEmphasis As Boolean
        //	<DispId(&H10D)> _
        //	Property AutoFormatAsYouTypeFormatListItemBeginning As Boolean
        //	<DispId(270)> _
        //	Property AutoFormatAsYouTypeDefineStyles As Boolean
        //	<DispId(&H10F)> _
        //	Property AutoFormatPlainTextWordMail As Boolean
        //	<DispId(&H110)> _
        //	Property AutoFormatAsYouTypeReplaceHyperlinks As Boolean
        //	<DispId(&H111)> _
        //	Property AutoFormatReplaceHyperlinks As Boolean
        //	<DispId(&H112)> _
        //	Property DefaultHighlightColorIndex As WdColorIndex
        //	<DispId(&H113)> _
        //	Property DefaultBorderLineStyle As WdLineStyle
        //	<DispId(&H114)> _
        //	Property CheckSpellingAsYouType As Boolean
        //	<DispId(&H115)> _
        //	Property CheckGrammarAsYouType As Boolean
        //	<DispId(&H116)> _
        //	Property IgnoreInternetAndFileAddresses As Boolean
        //	<DispId(&H117)> _
        //	Property ShowReadabilityStatistics As Boolean
        //	<DispId(280)> _
        //	Property IgnoreUppercase As Boolean
        //	<DispId(&H119)> _
        //	Property IgnoreMixedDigits As Boolean
        //	<DispId(&H11A)> _
        //	Property SuggestFromMainDictionaryOnly As Boolean
        //	<DispId(&H11B)> _
        //	Property SuggestSpellingCorrections As Boolean
        //	<DispId(&H11C)> _
        //	Property DefaultBorderLineWidth As WdLineWidth
        //	<DispId(&H11D)> _
        //	Property CheckGrammarWithSpelling As Boolean
        //	<DispId(&H11E)> _
        //	Property DefaultOpenFormat As WdOpenFormat
        //	<DispId(&H11F)> _
        //	Property PrintDraft As Boolean
        //	<DispId(&H120)> _
        //	Property PrintReverse As Boolean
        //	<DispId(&H121)> _
        //	Property MapPaperSize As Boolean
        //	<DispId(290)> _
        //	Property AutoFormatAsYouTypeApplyTables As Boolean
        //	<DispId(&H123)> _
        //	Property AutoFormatApplyFirstIndents As Boolean
        //	<DispId(&H126)> _
        //	Property AutoFormatMatchParentheses As Boolean
        //	<DispId(&H127)> _
        //	Property AutoFormatReplaceFarEastDashes As Boolean
        //	<DispId(&H128)> _
        //	Property AutoFormatDeleteAutoSpaces As Boolean
        //	<DispId(&H129)> _
        //	Property AutoFormatAsYouTypeApplyFirstIndents As Boolean
        //	<DispId(&H12A)> _
        //	Property AutoFormatAsYouTypeApplyDates As Boolean
        //	<DispId(&H12B)> _
        //	Property AutoFormatAsYouTypeApplyClosings As Boolean
        //	<DispId(300)> _
        //	Property AutoFormatAsYouTypeMatchParentheses As Boolean
        //	<DispId(&H12D)> _
        //	Property AutoFormatAsYouTypeReplaceFarEastDashes As Boolean
        //	<DispId(&H12E)> _
        //	Property AutoFormatAsYouTypeDeleteAutoSpaces As Boolean
        //	<DispId(&H12F)> _
        //	Property AutoFormatAsYouTypeInsertClosings As Boolean
        //	<DispId(&H130)> _
        //	Property AutoFormatAsYouTypeAutoLetterWizard As Boolean
        //	<DispId(&H131)> _
        //	Property AutoFormatAsYouTypeInsertOvers As Boolean
        //	<DispId(&H132)> _
        //	Property DisplayGridLines As Boolean
        //	<DispId(&H135)> _
        //	Property MatchFuzzyCase As Boolean
        //	<DispId(310)> _
        //	Property MatchFuzzyByte As Boolean
        //	<DispId(&H137)> _
        //	Property MatchFuzzyHiragana As Boolean
        //	<DispId(&H138)> _
        //	Property MatchFuzzySmallKana As Boolean
        //	<DispId(&H139)> _
        //	Property MatchFuzzyDash As Boolean
        //	<DispId(&H13A)> _
        //	Property MatchFuzzyIterationMark As Boolean
        //	<DispId(&H13B)> _
        //	Property MatchFuzzyKanji As Boolean
        //	<DispId(&H13C)> _
        //	Property MatchFuzzyOldKana As Boolean
        //	<DispId(&H13D)> _
        //	Property MatchFuzzyProlongedSoundMark As Boolean
        //	<DispId(&H13E)> _
        //	Property MatchFuzzyDZ As Boolean
        //	<DispId(&H13F)> _
        //	Property MatchFuzzyBV As Boolean
        //	<DispId(320)> _
        //	Property MatchFuzzyTC As Boolean
        //	<DispId(&H141)> _
        //	Property MatchFuzzyHF As Boolean
        //	<DispId(&H142)> _
        //	Property MatchFuzzyZJ As Boolean
        //	<DispId(&H143)> _
        //	Property MatchFuzzyAY As Boolean
        //	<DispId(&H144)> _
        //	Property MatchFuzzyKiKu As Boolean
        //	<DispId(&H145)> _
        //	Property MatchFuzzyPunctuation As Boolean
        //	<DispId(&H146)> _
        //	Property MatchFuzzySpace As Boolean
        //	<DispId(&H147)> _
        //	Property ApplyFarEastFontsToAscii As Boolean
        //	<DispId(&H148)> _
        //	Property ConvertHighAnsiToFarEast As Boolean
        //	<DispId(330)> _
        //	Property PrintOddPagesInAscendingOrder As Boolean
        //	<DispId(&H14B)> _
        //	Property PrintEvenPagesInAscendingOrder As Boolean
        //	<DispId(&H151)> _
        //	Property DefaultBorderColorIndex As WdColorIndex
        //	<DispId(&H152)> _
        //	Property EnableMisusedWordsDictionary As Boolean
        //	<DispId(&H153)> _
        //	Property AllowCombinedAuxiliaryForms As Boolean
        //	<DispId(340)> _
        //	Property HangulHanjaFastConversion As Boolean
        //	<DispId(&H155)> _
        //	Property CheckHangulEndings As Boolean
        //	<DispId(&H156)> _
        //	Property EnableHangulHanjaRecentOrdering As Boolean
        //	<DispId(&H157)> _
        //	Property MultipleWordConversionsMode As WdMultipleWordConversionsMode
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), TypeLibFunc(CShort(&H40)), DispId(&H14D)> _
        //Sub SetWPHelpOptions(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CommandKeyHelp As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional DocNavigationKeys As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional MouseSimulation As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional DemoGuidance As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional DemoSpeed As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional HelpType As Object)
        //	<DispId(&H158)> _
        //	Property DefaultBorderColor As WdColor
        //	<DispId(&H159)> _
        //	Property AllowPixelUnits As Boolean
        //	<DispId(&H15A)> _
        //	Property UseCharacterUnit As Boolean
        //	<DispId(&H15B)> _
        //	Property AllowCompoundNounProcessing As Boolean
        //	<DispId(&H18F)> _
        //	Property AutoKeyboardSwitching As Boolean
        //	<DispId(400)> _
        //	Property DocumentViewDirection As WdDocumentViewDirection
        //	<DispId(&H191)> _
        //	Property ArabicNumeral As WdArabicNumeral
        //	<DispId(&H192)> _
        //	Property MonthNames As WdMonthNames
        //	<DispId(&H193)> _
        //	Property CursorMovement As WdCursorMovement
        //	<DispId(&H194)> _
        //	Property VisualSelection As WdVisualSelection
        //	<DispId(&H195)> _
        //	Property ShowDiacritics As Boolean
        //	<DispId(&H196)> _
        //	Property ShowControlCharacters As Boolean
        //<DispId(&H197)> _
        //Property AddControlCharacters As Boolean
        //<DispId(&H197)> _
        //Function AddControlCharacters() As Boolean

        //	<DispId(&H198)> _
        //	Property AddBiDirectionalMarksWhenSavingTextFile As Boolean
        [PreserveSig]
        int AddBiDirectionalMarksWhenSavingTextFile_Get(out bool retVal);

        //	<DispId(&H199)> _
        //	Property StrictInitialAlefHamza As Boolean
        //	<DispId(410)> _
        //	Property StrictFinalYaa As Boolean
        //	<DispId(&H19B)> _
        //	Property HebrewMode As WdHebSpellStart
        //	<DispId(&H19C)> _
        //	Property ArabicMode As WdAraSpeller
        //	<DispId(&H19D)> _
        //	Property AllowClickAndTypeMouse As Boolean
        //	<DispId(&H19F)> _
        //	Property UseGermanSpellingReform As Boolean
        //	<DispId(&H1A2)> _
        //	Property InterpretHighAnsi As WdHighAnsiText
        //	<DispId(&H1A3)> _
        //	Property AddHebDoubleQuote As Boolean
        //	<DispId(420)> _
        //	Property UseDiffDiacColor As Boolean
        //	<DispId(&H1A5)> _
        //	Property DiacriticColorVal As WdColor
        //	<DispId(&H1A7)> _
        //	Property OptimizeForWord97byDefault As Boolean
        //	<DispId(&H1A8)> _
        //	Property LocalNetworkFile As Boolean
        //	<DispId(&H1A9)> _
        //	Property TypeNReplace As Boolean
        //	<DispId(&H1AA)> _
        //	Property SequenceCheck As Boolean
        //	<DispId(&H1AB)> _
        //	Property BackgroundOpen As Boolean
        //	<DispId(&H1AC)> _
        //	Property DisableFeaturesbyDefault As Boolean
        //	<DispId(&H1AD)> _
        //	Property PasteAdjustWordSpacing As Boolean
        //	<DispId(430)> _
        //	Property PasteAdjustParagraphSpacing As Boolean
        //	<DispId(&H1AF)> _
        //	Property PasteAdjustTableFormatting As Boolean
        //	<DispId(&H1B0)> _
        //	Property PasteSmartStyleBehavior As Boolean
        //	<DispId(&H1B1)> _
        //	Property PasteMergeFromPPT As Boolean
        //	<DispId(&H1B2)> _
        //	Property PasteMergeFromXL As Boolean
        //	<DispId(&H1B3)> _
        //	Property CtrlClickHyperlinkToOpen As Boolean
        //	<DispId(&H1B4)> _
        //	Property PictureWrapType As WdWrapTypeMerged
        //	<DispId(&H1B5)> _
        //	Property DisableFeaturesIntroducedAfterbyDefault As WdDisableFeaturesIntroducedAfter
        //	<DispId(&H1B6)> _
        //	Property PasteSmartCutPaste As Boolean
        //	<DispId(&H1B7)> _
        //	Property DisplayPasteOptions As Boolean
        //	<DispId(&H1B9)> _
        //	Property PromptUpdateStyle As Boolean
        //	<DispId(&H1BA)> _
        //	Property DefaultEPostageApp As <MarshalAs(UnmanagedType.BStr)> String
        //	<DispId(&H1BB)> _
        //	Property DefaultTextEncoding As MsoEncoding
        //	<DispId(&H1BC)> _
        //	Property LabelSmartTags As Boolean
        //	<DispId(&H1BD)> _
        //	Property DisplaySmartTagButtons As Boolean
        //	<DispId(&H1BE)> _
        //	Property WarnBeforeSavingPrintingSendingMarkup As Boolean
        //	<DispId(&H1BF)> _
        //	Property StoreRSIDOnSave As Boolean
        //	<DispId(&H1C0)> _
        //	Property ShowFormatError As Boolean
        //	<DispId(&H1C1)> _
        //	Property FormatScanning As Boolean
        //	<DispId(450)> _
        //	Property PasteMergeLists As Boolean
        //	<DispId(&H1C3)> _
        //	Property AutoCreateNewDrawings As Boolean
        //	<DispId(&H1C4)> _
        //	Property SmartParaSelection As Boolean
        //	<DispId(&H1C5)> _
        //	Property RevisionsBalloonPrintOrientation As WdRevisionsBalloonPrintOrientation
        //	<DispId(&H1C6)> _
        //	Property CommentsColor As WdColorIndex
        //	<DispId(&H1C7)> _
        //	Property PrintXMLTag As Boolean
        //	<DispId(&H1C8)> _
        //	Property PrintBackgrounds As Boolean
        //	<DispId(&H1C9)> _
        //	Property AllowReadingMode As Boolean
        //	<DispId(&H1CA)> _
        //	Property ShowMarkupOpenSave As Boolean
        //	<DispId(&H1CB)> _
        //	Property SmartCursoring As Boolean
        //	<DispId(460)> _
        //	Property MoveToTextMark As WdMoveToTextMark
        //	<DispId(&H1CD)> _
        //	Property MoveFromTextMark As WdMoveFromTextMark
        //	<DispId(&H1CE)> _
        //	Property BibliographyStyle As <MarshalAs(UnmanagedType.BStr)> String
        //	<DispId(&H1CF)> _
        //	Property BibliographySort As <MarshalAs(UnmanagedType.BStr)> String
        //	<DispId(&H1D0)> _
        //	Property InsertedCellColor As WdCellColor
        //	<DispId(&H1D1)> _
        //	Property DeletedCellColor As WdCellColor
        //	<DispId(&H1D2)> _
        //	Property MergedCellColor As WdCellColor
        //	<DispId(&H1D3)> _
        //	Property SplitCellColor As WdCellColor
        //	<DispId(&H1D4)> _
        //	Property ShowSelectionFloaties As Boolean
        //	<DispId(&H1D5)> _
        //	Property ShowMenuFloaties As Boolean
        //	<DispId(470)> _
        //	Property ShowDevTools As Boolean
        //	<DispId(&H1D8)> _
        //	Property EnableLivePreview As Boolean
        //	<DispId(&H1DA)> _
        //	Property OMathAutoBuildUp As Boolean
        //	<DispId(&H1DC)> _
        //	Property AlwaysUseClearType As Boolean
        //	<DispId(&H1DD)> _
        //	Property PasteFormatWithinDocument As WdPasteOptions
        //	<DispId(&H1DE)> _
        //	Property PasteFormatBetweenDocuments As WdPasteOptions
        //	<DispId(&H1DF)> _
        //	Property PasteFormatBetweenStyledDocuments As WdPasteOptions
        //	<DispId(480)> _
        //	Property PasteFormatFromExternalSource As WdPasteOptions
        //	<DispId(&H1E1)> _
        //	Property PasteOptionKeepBulletsAndNumbers As Boolean
        //	<DispId(&H1E2)> _
        //	Property INSKeyForOvertype As Boolean
        //	<DispId(&H1E3)> _
        //	Property RepeatWord As Boolean
        //	<DispId(&H1E4)> _
        //	Property FrenchReform As WdFrenchSpeller
        //	<DispId(&H1E5)> _
        //	Property ContextualSpeller As Boolean
        //	<DispId(&H1E6)> _
        //	Property MoveToTextColor As WdColorIndex
        //	<DispId(&H1E7)> _
        //	Property MoveFromTextColor As WdColorIndex
        //	<DispId(&H1E8)> _
        //	Property OMathCopyLF As Boolean
        //	<DispId(&H1E9)> _
        //	Property UseNormalStyleForList As Boolean
        //	<DispId(490)> _
        //	Property AllowOpenInDraftView As Boolean
        //	<DispId(&H1EC)> _
        //	Property EnableLegacyIMEMode As Boolean
        //	<DispId(&H1ED)> _
        //	Property DoNotPromptForConvert As Boolean
        //	<DispId(&H1EE)> _
        //	Property PrecisePositioning As Boolean
    }
}