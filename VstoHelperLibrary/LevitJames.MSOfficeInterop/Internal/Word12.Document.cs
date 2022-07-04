// © Copyright 2018 Levit & James, Inc.

using System.ComponentModel;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace LevitJames.MSOffice.Internal
{
    [ComImport]
    [TypeLibType(flags: 0x1050)]
    [Guid("0002096B-0000-0000-C000-000000000046")]
    [DefaultMember("Name")]
    internal interface Document12
    {
        //7 to get past the IDispatch members
        //The TrackFormatting member is member number 359 so that's 359-7 gives us 352 hence _VtblGap7_352
        //Note: REMEMBER that read/write properties take 2 vtable slots.
        [MethodImpl(MethodCodeType = MethodCodeType.Runtime)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        void _VtblGap7_154();

        //<DispId(0)> _
        //Default ReadOnly Property Name As <MarshalAs(UnmanagedType.BStr)> String
        //<DispId(1)> _
        //ReadOnly Property Application As <MarshalAs(UnmanagedType.Interface)> Application
        //<DispId(&H3E9)> _
        //ReadOnly Property Creator As Integer
        //<DispId(&H3EA)> _
        //ReadOnly Property Parent As <MarshalAs(UnmanagedType.IDispatch)> Object
        //<DispId(&H3E8)> _
        //ReadOnly Property BuiltInDocumentProperties As <MarshalAs(UnmanagedType.IDispatch)> Object
        //<DispId(2)> _
        //ReadOnly Property CustomDocumentProperties As <MarshalAs(UnmanagedType.IDispatch)> Object
        //<DispId(3)> _
        //ReadOnly Property Path As <MarshalAs(UnmanagedType.BStr)> String
        //<DispId(4)> _
        //ReadOnly Property Bookmarks As <MarshalAs(UnmanagedType.Interface)> Bookmarks
        //<DispId(6)> _
        //ReadOnly Property Tables As <MarshalAs(UnmanagedType.Interface)> Tables
        //<DispId(7)> _
        //ReadOnly Property Footnotes As <MarshalAs(UnmanagedType.Interface)> Footnotes
        //<DispId(8)> _
        //ReadOnly Property Endnotes As <MarshalAs(UnmanagedType.Interface)> Endnotes
        //<DispId(9)> _
        //ReadOnly Property Comments As <MarshalAs(UnmanagedType.Interface)> Comments
        //<DispId(10)> _
        //ReadOnly Property Type As WdDocumentType
        //<DispId(11)> _
        //Property AutoHyphenation As Boolean
        //<DispId(12)> _
        //Property HyphenateCaps As Boolean
        //<DispId(13)> _
        //Property HyphenationZone As Integer
        //<DispId(14)> _
        //Property ConsecutiveHyphensLimit As Integer
        //<DispId(15)> _
        //ReadOnly Property Sections As <MarshalAs(UnmanagedType.Interface)> Sections
        //<DispId(&H10)> _
        //ReadOnly Property Paragraphs As <MarshalAs(UnmanagedType.Interface)> Paragraphs
        //<DispId(&H11)> _
        //ReadOnly Property Words As <MarshalAs(UnmanagedType.Interface)> Words
        //<DispId(&H12)> _
        //ReadOnly Property Sentences As <MarshalAs(UnmanagedType.Interface)> Sentences
        //<DispId(&H13)> _
        //ReadOnly Property Characters As <MarshalAs(UnmanagedType.Interface)> Characters
        //<DispId(20)> _
        //ReadOnly Property Fields As <MarshalAs(UnmanagedType.Interface)> Fields
        //<DispId(&H15)> _
        //ReadOnly Property FormFields As <MarshalAs(UnmanagedType.Interface)> FormFields
        //<DispId(&H16)> _
        //ReadOnly Property Styles As <MarshalAs(UnmanagedType.Interface)> Styles
        //<DispId(&H17)> _
        //ReadOnly Property Frames As <MarshalAs(UnmanagedType.Interface)> Frames
        //<DispId(&H19)> _
        //ReadOnly Property TablesOfFigures As <MarshalAs(UnmanagedType.Interface)> TablesOfFigures
        //<DispId(&H1A)> _
        //ReadOnly Property Variables As <MarshalAs(UnmanagedType.Interface)> Variables
        //<DispId(&H1B)> _
        //ReadOnly Property MailMerge As <MarshalAs(UnmanagedType.Interface)> MailMerge
        //<DispId(&H1C)> _
        //ReadOnly Property Envelope As <MarshalAs(UnmanagedType.Interface)> Envelope
        //<DispId(&H1D)> _
        //ReadOnly Property FullName As <MarshalAs(UnmanagedType.BStr)> String
        //<DispId(30)> _
        //ReadOnly Property Revisions As <MarshalAs(UnmanagedType.Interface)> Revisions
        //<DispId(&H1F)> _
        //ReadOnly Property TablesOfContents As <MarshalAs(UnmanagedType.Interface)> TablesOfContents
        //<DispId(&H20)> _
        //ReadOnly Property TablesOfAuthorities As <MarshalAs(UnmanagedType.Interface)> TablesOfAuthorities
        //<DispId(&H44D)> _
        //Property PageSetup As <MarshalAs(UnmanagedType.Interface)> PageSetup
        //<DispId(&H22)> _
        //ReadOnly Property Windows As <MarshalAs(UnmanagedType.Interface)> Windows
        //<DispId(&H23)> _
        //Property HasRoutingSlip As Boolean
        //<DispId(&H24)> _
        //ReadOnly Property RoutingSlip As <MarshalAs(UnmanagedType.Interface)> RoutingSlip
        //<DispId(&H25)> _
        //ReadOnly Property Routed As Boolean
        //<DispId(&H26)> _
        //ReadOnly Property TablesOfAuthoritiesCategories As <MarshalAs(UnmanagedType.Interface)> TablesOfAuthoritiesCategories
        //<DispId(&H27)> _
        //ReadOnly Property Indexes As <MarshalAs(UnmanagedType.Interface)> Indexes
        //<DispId(40)> _
        //Property Saved As Boolean
        //<DispId(&H29)> _
        //ReadOnly Property Content As <MarshalAs(UnmanagedType.Interface)> Range
        //<DispId(&H2A)> _
        //ReadOnly Property ActiveWindow As <MarshalAs(UnmanagedType.Interface)> Window
        //<DispId(&H2B)> _
        //Property Kind As WdDocumentKind
        //<DispId(&H2C)> _
        //ReadOnly Property [ReadOnly] As Boolean
        //<DispId(&H2D)> _
        //ReadOnly Property Subdocuments As <MarshalAs(UnmanagedType.Interface)> Subdocuments
        //<DispId(&H2E)> _
        //ReadOnly Property IsMasterDocument As Boolean
        //<DispId(&H30)> _
        //Property DefaultTabStop As Single
        //<DispId(50)> _
        //Property EmbedTrueTypeFonts As Boolean
        //<DispId(&H33)> _
        //Property SaveFormsData As Boolean
        //<DispId(&H34)> _
        //Property ReadOnlyRecommended As Boolean
        //<DispId(&H35)> _
        //Property SaveSubsetFonts As Boolean
        //<DispId(&H37)> _
        //Property Compatibility(ByVal Type As WdCompatibility) As Boolean
        //<DispId(&H38)> _
        //ReadOnly Property StoryRanges As <MarshalAs(UnmanagedType.Interface)> StoryRanges
        //<DispId(&H39)> _
        //ReadOnly Property CommandBars As <MarshalAs(UnmanagedType.Interface)> CommandBars
        //<DispId(&H3A)> _
        //ReadOnly Property IsSubdocument As Boolean
        //<DispId(&H3B)> _
        //ReadOnly Property SaveFormat As Integer
        //<DispId(60)> _
        //ReadOnly Property ProtectionType As WdProtectionType
        //<DispId(&H3D)> _
        //ReadOnly Property Hyperlinks As <MarshalAs(UnmanagedType.Interface)> Hyperlinks
        //<DispId(&H3E)> _
        //ReadOnly Property Shapes As <MarshalAs(UnmanagedType.Interface)> Shapes
        //<DispId(&H3F)> _
        //ReadOnly Property ListTemplates As <MarshalAs(UnmanagedType.Interface)> ListTemplates
        //<DispId(&H40)> _
        //ReadOnly Property Lists As <MarshalAs(UnmanagedType.Interface)> Lists
        //<DispId(&H42)> _
        //Property UpdateStylesOnOpen As Boolean
        //<DispId(&H43)> _
        //Property AttachedTemplate As <MarshalAs(UnmanagedType.Struct)> Object
        //<DispId(&H44)> _
        //ReadOnly Property InlineShapes As <MarshalAs(UnmanagedType.Interface)> InlineShapes
        //<DispId(&H45)> _
        //Property Background As <MarshalAs(UnmanagedType.Interface)> Shape
        //<DispId(70)> _
        //Property GrammarChecked As Boolean
        //<DispId(&H47)> _
        //Property SpellingChecked As Boolean
        //<DispId(&H48)> _
        //Property ShowGrammaticalErrors As Boolean
        //<DispId(&H49)> _
        //Property ShowSpellingErrors As Boolean
        //<DispId(&H4B)> _
        //ReadOnly Property Versions As <MarshalAs(UnmanagedType.Interface)> Versions
        //<DispId(&H4C)> _
        //Property ShowSummary As Boolean
        //<DispId(&H4D)> _
        //Property SummaryViewMode As WdSummaryMode
        //<DispId(&H4E)> _
        //Property SummaryLength As Integer
        //<DispId(&H4F)> _
        //Property PrintFractionalWidths As Boolean
        //<DispId(80)> _
        //Property PrintPostScriptOverText As Boolean
        //<DispId(&H52)> _
        //ReadOnly Property Container As <MarshalAs(UnmanagedType.IDispatch)> Object
        //<DispId(&H53)> _
        //Property PrintFormsData As Boolean
        //<DispId(&H54)> _
        //ReadOnly Property ListParagraphs As <MarshalAs(UnmanagedType.Interface)> ListParagraphs
        //<DispId(&H55)> _
        //WriteOnly Property Password As String
        //<DispId(&H56)> _
        //WriteOnly Property WritePassword As String
        //<DispId(&H57)> _
        //ReadOnly Property HasPassword As Boolean
        //<DispId(&H58)> _
        //ReadOnly Property WriteReserved As Boolean
        //<DispId(90)> _
        //Property ActiveWritingStyle(ByRef LanguageID As Object) As <MarshalAs(UnmanagedType.BStr)> String
        //<DispId(&H5C)> _
        //Property UserControl As Boolean
        //<DispId(&H5D)> _
        //Property HasMailer As Boolean
        //<DispId(&H5E)> _
        //ReadOnly Property Mailer As <MarshalAs(UnmanagedType.Interface)> Mailer
        //<DispId(&H60)> _
        //ReadOnly Property ReadabilityStatistics As <MarshalAs(UnmanagedType.Interface)> ReadabilityStatistics
        //<DispId(&H61)> _
        //ReadOnly Property GrammaticalErrors As <MarshalAs(UnmanagedType.Interface)> ProofreadingErrors
        //<DispId(&H62)> _
        //ReadOnly Property SpellingErrors As <MarshalAs(UnmanagedType.Interface)> ProofreadingErrors
        //<DispId(&H63)> _
        //ReadOnly Property VBProject As <MarshalAs(UnmanagedType.Interface)> VBProject
        //<DispId(100)> _
        //ReadOnly Property FormsDesign As Boolean
        //<DispId(-2147418112)> _
        //Property _CodeName As <MarshalAs(UnmanagedType.BStr)> String
        //<DispId(&H106)> _
        //ReadOnly Property CodeName As <MarshalAs(UnmanagedType.BStr)> String
        //<DispId(300)> _
        //Property SnapToGrid As Boolean
        //<DispId(&H12D)> _
        //Property SnapToShapes As Boolean
        //<DispId(&H12E)> _
        //Property GridDistanceHorizontal As Single
        //<DispId(&H12F)> _
        //Property GridDistanceVertical As Single
        //<DispId(&H130)> _
        //Property GridOriginHorizontal As Single
        //<DispId(&H131)> _
        //Property GridOriginVertical As Single
        //<DispId(&H132)> _
        //Property GridSpaceBetweenHorizontalLines As Integer
        //<DispId(&H133)> _
        //Property GridSpaceBetweenVerticalLines As Integer
        //<DispId(&H134)> _
        //Property GridOriginFromMargin As Boolean
        //<DispId(&H135)> _
        //Property KerningByAlgorithm As Boolean
        //<DispId(310)> _
        //Property JustificationMode As WdJustificationMode
        //<DispId(&H137)> _
        //Property FarEastLineBreakLevel As WdFarEastLineBreakLevel
        //<DispId(&H138)> _
        //Property NoLineBreakBefore As <MarshalAs(UnmanagedType.BStr)> String
        //<DispId(&H139)> _
        //Property NoLineBreakAfter As <MarshalAs(UnmanagedType.BStr)> String
        //<DispId(&H13A)> _
        //Property TrackRevisions As Boolean
        [PreserveSig]
        int TrackRevisions_Get(out bool retVal); //161

        [PreserveSig]
        int TrackRevisions_Let(bool sr); //162

        [MethodImpl(MethodCodeType = MethodCodeType.Runtime)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        void _VtblGap_196(); //359-163=196

        //<DispId(&H13B)> _
        //Property PrintRevisions As Boolean
        //<DispId(&H13C)> _
        //Property ShowRevisions As Boolean
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H451)> _
        //Sub Close(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional SaveChanges As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional OriginalFormat As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional RouteDocument As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H66), TypeLibFunc(CShort(&H40))> _
        //Sub SaveAs2000(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional FileName As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional FileFormat As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional LockComments As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Password As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional AddToRecentFiles As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional WritePassword As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ReadOnlyRecommended As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional EmbedTrueTypeFonts As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional SaveNativePictureFormat As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional SaveFormsData As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional SaveAsAOCELetter As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H67)> _
        //Sub Repaginate()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H68)> _
        //Sub FitToPages()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H69)> _
        //Sub ManualHyphenation()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&HFFFF)> _
        //Sub [Select]()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H6A)> _
        //Sub DataForm()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H6B), TypeLibFunc(CShort(&H40))> _
        //Sub Route()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H6C)> _
        //Sub Save()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H6D), TypeLibFunc(CShort(&H40))> _
        //Sub PrintOutOld(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Background As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Append As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Range As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional OutputFileName As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional From As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional [To] As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Item As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Copies As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Pages As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PageType As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintToFile As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Collate As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ActivePrinterMacGX As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ManualDuplexPrint As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(110)> _
        //Sub SendMail()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H7D0)> _
        //Function Range(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Start As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional [End] As Object) As <MarshalAs(UnmanagedType.Interface)> Range
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H70)> _
        //Sub RunAutoMacro(<[In]()> ByVal Which As WdAutoMacros)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H71)> _
        //Sub Activate()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H72)> _
        //Sub PrintPreview()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H73)> _
        //Function [GoTo](<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional What As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Which As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Count As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Name As Object) As <MarshalAs(UnmanagedType.Interface)> Range
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H74)> _
        //Function Undo(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Times As Object) As Boolean
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H75)> _
        //Function Redo(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Times As Object) As Boolean
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H76)> _
        //Function ComputeStatistics(<[In]> ByVal Statistic As WdStatistic, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional IncludeFootnotesAndEndnotes As Object) As Integer
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H77)> _
        //Sub MakeCompatibilityDefault()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), TypeLibFunc(CShort(&H40)), DispId(120)> _
        //Sub Protect2002(<[In]> ByVal Type As WdProtectionType, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional NoReset As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Password As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H79)> _
        //Sub Unprotect(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Password As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H7A)> _
        //Sub EditionOptions(<[In]> ByVal Type As WdEditionType, <[In]> ByVal [Option] As WdEditionOption, <[In], MarshalAs(UnmanagedType.BStr)> ByVal Name As String, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Format As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H7B)> _
        //Sub RunLetterWizard(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional LetterContent As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional WizardMode As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H7C)> _
        //Function GetLetterContent() As <MarshalAs(UnmanagedType.Interface)> LetterContent
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H7D)> _
        //Sub SetLetterContent(<[In](), MarshalAs(UnmanagedType.Struct)> ByRef LetterContent As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H7E)> _
        //Sub CopyStylesFromTemplate(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal Template As String)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H7F)> _
        //Sub UpdateStyles()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H83)> _
        //Sub CheckGrammar()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H84)> _
        //Sub CheckSpelling(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional IgnoreUppercase As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional AlwaysSuggest As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary2 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary3 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary4 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary5 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary6 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary7 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary8 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary9 As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CustomDictionary10 As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H87)> _
        //Sub FollowHyperlink(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Address As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional SubAddress As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional NewWindow As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional AddHistory As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ExtraInfo As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Method As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional HeaderInfo As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H88)> _
        //Sub AddToFavorites()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H89)> _
        //Sub Reload()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H8A)> _
        //Function AutoSummarize(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Length As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Mode As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional UpdateProperties As Object) As <MarshalAs(UnmanagedType.Interface)> Range
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(140)> _
        //Sub RemoveNumbers(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional NumberType As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H8D)> _
        //Sub ConvertNumbersToText(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional NumberType As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H8E)> _
        //Function CountNumberedItems(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional NumberType As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Level As Object) As Integer
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H8F)> _
        //Sub Post()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H90)> _
        //Sub ToggleFormsDesign()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H91), TypeLibFunc(CShort(&H40))> _
        //Sub Compare2000(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal Name As String)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H92)> _
        //Sub UpdateSummaryProperties()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H93)> _
        //Function GetCrossReferenceItems(<[In](), MarshalAs(UnmanagedType.Struct)> ByRef ReferenceType As Object) As <MarshalAs(UnmanagedType.Struct)> Object
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H94)> _
        //Sub AutoFormat()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H95)> _
        //Sub ViewCode()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(150)> _
        //Sub ViewPropertyBrowser()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(250)> _
        //Sub ForwardMailer()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&HFB)> _
        //Sub Reply()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&HFC)> _
        //Sub ReplyAll()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&HFD)> _
        //Sub SendMailer(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional FileFormat As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Priority As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&HFE)> _
        //Sub UndoClear()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&HFF)> _
        //Sub PresentIt()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H100)> _
        //Sub SendFax(<[In], MarshalAs(UnmanagedType.BStr)> ByVal Address As String, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Subject As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), TypeLibFunc(CShort(&H40)), DispId(&H101)> _
        //Sub Merge2000(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal FileName As String)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H102)> _
        //Sub ClosePrintPreview()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H103)> _
        //Sub CheckConsistency()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(260)> _
        //Function CreateLetterContent(<[In], MarshalAs(UnmanagedType.BStr)> ByVal DateFormat As String, <[In]> ByVal IncludeHeaderFooter As Boolean, <[In], MarshalAs(UnmanagedType.BStr)> ByVal PageDesign As String, <[In]> ByVal LetterStyle As WdLetterStyle, <[In]> ByVal Letterhead As Boolean, <[In]> ByVal LetterheadLocation As WdLetterheadLocation, <[In]> ByVal LetterheadSize As Single, <[In], MarshalAs(UnmanagedType.BStr)> ByVal RecipientName As String, <[In], MarshalAs(UnmanagedType.BStr)> ByVal RecipientAddress As String, <[In], MarshalAs(UnmanagedType.BStr)> ByVal Salutation As String, <[In]> ByVal SalutationType As WdSalutationType, <[In], MarshalAs(UnmanagedType.BStr)> ByVal RecipientReference As String, <[In], MarshalAs(UnmanagedType.BStr)> ByVal MailingInstructions As String, <[In], MarshalAs(UnmanagedType.BStr)> ByVal AttentionLine As String, <[In], MarshalAs(UnmanagedType.BStr)> ByVal Subject As String, <[In], MarshalAs(UnmanagedType.BStr)> ByVal CCList As String, <[In], MarshalAs(UnmanagedType.BStr)> ByVal ReturnAddress As String, <[In], MarshalAs(UnmanagedType.BStr)> ByVal SenderName As String, <[In], MarshalAs(UnmanagedType.BStr)> ByVal Closing As String, <[In], MarshalAs(UnmanagedType.BStr)> ByVal SenderCompany As String, <[In], MarshalAs(UnmanagedType.BStr)> ByVal SenderJobTitle As String, <[In], MarshalAs(UnmanagedType.BStr)> ByVal SenderInitials As String, <[In]> ByVal EnclosureNumber As Integer, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional InfoBlock As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional RecipientCode As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional RecipientGender As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ReturnAddressShortForm As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional SenderCity As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional SenderCode As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional SenderGender As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional SenderReference As Object) As <MarshalAs(UnmanagedType.Interface)> LetterContent
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H13D)> _
        //Sub AcceptAllRevisions()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H13E)> _
        //Sub RejectAllRevisions()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H97)> _
        //Sub DetectLanguage()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H142)> _
        //Sub ApplyTheme(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal Name As String)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H143)> _
        //Sub RemoveTheme()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H145)> _
        //Sub WebPagePreview()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H14B)> _
        //Sub ReloadAs(<[In]()> ByVal Encoding As MsoEncoding)
        //<DispId(540)> _
        //ReadOnly Property ActiveTheme As <MarshalAs(UnmanagedType.BStr)> String
        //<DispId(&H21D)> _
        //ReadOnly Property ActiveThemeDisplayName As <MarshalAs(UnmanagedType.BStr)> String
        //<DispId(&H13F)> _
        //ReadOnly Property Email As <MarshalAs(UnmanagedType.Interface)> Email
        //<DispId(320)> _
        //ReadOnly Property Scripts As <MarshalAs(UnmanagedType.Interface)> Scripts
        //<DispId(&H141)> _
        //Property LanguageDetected As Boolean
        //<DispId(&H146)> _
        //Property FarEastLineBreakLanguage As WdFarEastLineBreakLanguageID
        //<DispId(&H147)> _
        //ReadOnly Property Frameset As <MarshalAs(UnmanagedType.Interface)> Frameset
        //<DispId(&H148)> _
        //Property ClickAndTypeParagraphStyle As <MarshalAs(UnmanagedType.Struct)> Object
        //<DispId(&H149)> _
        //ReadOnly Property HTMLProject As <MarshalAs(UnmanagedType.Interface)> HTMLProject
        //<DispId(330)> _
        //ReadOnly Property WebOptions As <MarshalAs(UnmanagedType.Interface)> WebOptions
        //<DispId(&H14C)> _
        //ReadOnly Property OpenEncoding As MsoEncoding
        //<DispId(&H14D)> _
        //Property SaveEncoding As MsoEncoding
        //<DispId(&H14E)> _
        //Property OptimizeForWord97 As Boolean
        //<DispId(&H14F)> _
        //ReadOnly Property VBASigned As Boolean
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), TypeLibFunc(CShort(&H40)), DispId(&H1BC)> _
        //Sub PrintOut2000(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Background As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Append As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Range As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional OutputFileName As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional From As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional [To] As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Item As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Copies As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Pages As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PageType As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintToFile As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Collate As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ActivePrinterMacGX As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ManualDuplexPrint As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomColumn As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomRow As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomPaperWidth As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomPaperHeight As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1BD), TypeLibFunc(CShort(&H40))> _
        //Sub sblt(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal s As String)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1BF)> _
        //Sub ConvertVietDoc(<[In]()> ByVal CodePageOrigin As Integer)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1BE)> _
        //Sub PrintOut(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Background As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Append As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Range As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional OutputFileName As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional From As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional [To] As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Item As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Copies As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Pages As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PageType As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintToFile As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Collate As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ActivePrinterMacGX As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ManualDuplexPrint As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomColumn As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomRow As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomPaperWidth As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomPaperHeight As Object)
        //<DispId(&H150)> _
        //ReadOnly Property MailEnvelope As <MarshalAs(UnmanagedType.Interface)> MsoEnvelope
        //<DispId(&H151)> _
        //Property DisableFeatures As Boolean
        //<DispId(&H152)> _
        //Property DoNotEmbedSystemFonts As Boolean
        //<DispId(&H153)> _
        //ReadOnly Property Signatures As <MarshalAs(UnmanagedType.Interface)> SignatureSet
        //<DispId(340)> _
        //Property DefaultTargetFrame As <MarshalAs(UnmanagedType.BStr)> String
        //<DispId(&H156)> _
        //ReadOnly Property HTMLDivisions As <MarshalAs(UnmanagedType.Interface)> HTMLDivisions
        //<DispId(&H157)> _
        //Property DisableFeaturesIntroducedAfter As WdDisableFeaturesIntroducedAfter
        //<DispId(&H158)> _
        //Property RemovePersonalInformation As Boolean
        //<DispId(&H15A)> _
        //ReadOnly Property SmartTags As <MarshalAs(UnmanagedType.Interface)> SmartTags
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), TypeLibFunc(CShort(&H40)), DispId(&H159)> _
        //Sub Compare2002(<[In], MarshalAs(UnmanagedType.BStr)> ByVal Name As String, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional AuthorName As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CompareTarget As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional DetectFormatChanges As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional IgnoreAllComparisonWarnings As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional AddToRecentFiles As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H15D)> _
        //Sub CheckIn(<[In]> ByVal Optional SaveChanges As Boolean = True, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Comments As Object, <[In]> ByVal Optional MakePublic As Boolean = False)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H15F)> _
        //Function CanCheckin() As Boolean
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H16A)> _
        //Sub Merge(<[In], MarshalAs(UnmanagedType.BStr)> ByVal FileName As String, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional MergeTarget As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional DetectFormatChanges As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional UseFormattingFrom As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional AddToRecentFiles As Object)
        //<DispId(&H15B)> _
        //Property EmbedSmartTags As Boolean
        //<DispId(&H15C)> _
        //Property SmartTagsAsXMLProps As Boolean
        //<DispId(&H165)> _
        //Property TextEncoding As MsoEncoding
        //<DispId(&H166)> _
        //Property TextLineEnding As WdLineEndingType
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H161)> _
        //Sub SendForReview(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Recipients As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Subject As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ShowMessage As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional IncludeAttachment As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H162)> _
        //Sub ReplyWithChanges(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ShowMessage As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H164)> _
        //Sub EndReview()
        //<DispId(360)> _
        //ReadOnly Property StyleSheets As <MarshalAs(UnmanagedType.Interface)> StyleSheets
        //<DispId(&H16D)> _
        //ReadOnly Property DefaultTableStyle As <MarshalAs(UnmanagedType.Struct)> Object
        //<DispId(&H16F)> _
        //ReadOnly Property PasswordEncryptionProvider As <MarshalAs(UnmanagedType.BStr)> String
        //<DispId(&H170)> _
        //ReadOnly Property PasswordEncryptionAlgorithm As <MarshalAs(UnmanagedType.BStr)> String
        //<DispId(&H171)> _
        //ReadOnly Property PasswordEncryptionKeyLength As Integer
        //<DispId(370)> _
        //ReadOnly Property PasswordEncryptionFileProperties As Boolean
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H169)> _
        //Sub SetPasswordEncryptionOptions(<[In], MarshalAs(UnmanagedType.BStr)> ByVal PasswordEncryptionProvider As String, <[In], MarshalAs(UnmanagedType.BStr)> ByVal PasswordEncryptionAlgorithm As String, <[In]> ByVal PasswordEncryptionKeyLength As Integer, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PasswordEncryptionFileProperties As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H16B)> _
        //Sub RecheckSmartTags()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H16C)> _
        //Sub RemoveSmartTags()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H16E)> _
        //Sub SetDefaultTableStyle(<[In](), MarshalAs(UnmanagedType.Struct)> ByRef Style As Object, <[In]()> ByVal SetInTemplate As Boolean)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H173)> _
        //Sub DeleteAllComments()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H174)> _
        //Sub AcceptAllRevisionsShown()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H175)> _
        //Sub RejectAllRevisionsShown()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H176)> _
        //Sub DeleteAllCommentsShown()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H177)> _
        //Sub ResetFormFields()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H178)> _
        //Sub SaveAs(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional FileName As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional FileFormat As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional LockComments As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Password As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional AddToRecentFiles As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional WritePassword As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ReadOnlyRecommended As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional EmbedTrueTypeFonts As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional SaveNativePictureFormat As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional SaveFormsData As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional SaveAsAOCELetter As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Encoding As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional InsertLineBreaks As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional AllowSubstitutions As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional LineEnding As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional AddBiDiMarks As Object)
        //<DispId(&H179)> _
        //Property EmbedLinguisticData As Boolean
        //<DispId(&H1C0)> _
        //Property FormattingShowFont As Boolean
        //<DispId(&H1C1)> _
        //Property FormattingShowClear As Boolean
        //<DispId(450)> _
        //Property FormattingShowParagraph As Boolean
        //<DispId(&H1C3)> _
        //Property FormattingShowNumbering As Boolean
        //<DispId(&H1C4)> _
        //Property FormattingShowFilter As WdShowFilter
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H17A)> _
        //Sub CheckNewSmartTags()
        //<DispId(&H1C5)> _
        //ReadOnly Property Permission As <MarshalAs(UnmanagedType.Interface)> Permission
        //<DispId(460)> _
        //ReadOnly Property XMLNodes As <MarshalAs(UnmanagedType.Interface)> XMLNodes
        //<DispId(&H1CD)> _
        //ReadOnly Property XMLSchemaReferences As <MarshalAs(UnmanagedType.Interface)> XMLSchemaReferences
        //<DispId(&H1CE)> _
        //ReadOnly Property SmartDocument As <MarshalAs(UnmanagedType.Interface)> SmartDocument
        //<DispId(&H1CF)> _
        //ReadOnly Property SharedWorkspace As <MarshalAs(UnmanagedType.Interface)> SharedWorkspace
        //<DispId(&H1D2)> _
        //ReadOnly Property Sync As <MarshalAs(UnmanagedType.Interface)> Sync
        //<DispId(&H1D7)> _
        //Property EnforceStyle As Boolean
        //<DispId(&H1D8)> _
        //Property AutoFormatOverride As Boolean
        //<DispId(&H1D9)> _
        //Property XMLSaveDataOnly As Boolean
        //<DispId(&H1DD)> _
        //Property XMLHideNamespaces As Boolean
        //<DispId(&H1DE)> _
        //Property XMLShowAdvancedErrors As Boolean
        //<DispId(&H1DA)> _
        //Property XMLUseXSLTWhenSaving As Boolean
        //<DispId(&H1DB)> _
        //Property XMLSaveThroughXSLT As <MarshalAs(UnmanagedType.BStr)> String
        //<DispId(&H1DC)> _
        //ReadOnly Property DocumentLibraryVersions As <MarshalAs(UnmanagedType.Interface)> DocumentLibraryVersions
        //<DispId(&H1E1)> _
        //Property ReadingModeLayoutFrozen As Boolean
        //<DispId(&H1E4)> _
        //Property RemoveDateAndTime As Boolean
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1D0)> _
        //Sub SendFaxOverInternet(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Recipients As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Subject As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ShowMessage As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(500)> _
        //Sub TransformDocument(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal Path As String, <[In]()> Optional ByVal DataOnly As Boolean = True)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1D3)> _
        //Sub Protect(<[In]> ByVal Type As WdProtectionType, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional NoReset As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Password As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional UseIRM As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional EnforceStyleLock As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1D4)> _
        //Sub SelectAllEditableRanges(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional EditorID As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1D5)> _
        //Sub DeleteAllEditableRanges(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional EditorID As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1DF)> _
        //Sub DeleteAllInkAnnotations()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), TypeLibFunc(CShort(&H40)), DispId(&H1E2)> _
        //Sub AddDocumentWorkspaceHeader(<[In]()> ByVal RichFormat As Boolean, <[In](), MarshalAs(UnmanagedType.BStr)> ByVal Url As String, <[In](), MarshalAs(UnmanagedType.BStr)> ByVal Title As String, <[In](), MarshalAs(UnmanagedType.BStr)> ByVal Description As String, <[In](), MarshalAs(UnmanagedType.BStr)> ByVal ID As String)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1E3), TypeLibFunc(CShort(&H40))> _
        //Sub RemoveDocumentWorkspaceHeader(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal ID As String)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1E5)> _
        //Sub [Compare](<[In], MarshalAs(UnmanagedType.BStr)> ByVal Name As String, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional AuthorName As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional CompareTarget As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional DetectFormatChanges As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional IgnoreAllComparisonWarnings As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional AddToRecentFiles As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional RemovePersonalInformation As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional RemoveDateAndTime As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1E7)> _
        //Sub RemoveLockedStyles()
        //<DispId(&H1E6)> _
        //ReadOnly Property ChildNodeSuggestions As <MarshalAs(UnmanagedType.Interface)> XMLChildNodeSuggestions
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1E8)> _
        //Function SelectSingleNode(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal XPath As String, <[In](), MarshalAs(UnmanagedType.BStr)> Optional ByVal PrefixMapping As String = "", <[In]()> Optional ByVal FastSearchSkippingTextNodes As Boolean = True) As <MarshalAs(UnmanagedType.Interface)> XMLNode
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1E9)> _
        //Function SelectNodes(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal XPath As String, <[In](), MarshalAs(UnmanagedType.BStr)> Optional ByVal PrefixMapping As String = "", <[In]()> Optional ByVal FastSearchSkippingTextNodes As Boolean = True) As <MarshalAs(UnmanagedType.Interface)> XMLNodes
        //<DispId(490)> _
        //ReadOnly Property XMLSchemaViolations As <MarshalAs(UnmanagedType.Interface)> XMLNodes
        //<DispId(&H1EB)> _
        //Property ReadingLayoutSizeX As Integer
        //<DispId(&H1EC)> _
        //Property ReadingLayoutSizeY As Integer
        //<DispId(&H1ED)> _
        //Property StyleSortMethod As WdStyleSort
        //<DispId(&H1F0)> _
        //ReadOnly Property ContentTypeProperties As <MarshalAs(UnmanagedType.Interface)> MetaProperties
        //<DispId(&H1F3)> _
        //Property TrackMoves As Boolean
        [PreserveSig]
        int TrackFormatting_Get(out bool retVal); //359

        [PreserveSig]
        int TrackFormatting_Let(bool sr); //360


        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [TypeLibFunc(flags: 0x40)]
        [DispId(dispId: 0x1F7)]
        void Dummy1();

        [DispId(dispId: 0x1F8)]
        OMaths OMaths { get; }

        [MethodImpl(MethodCodeType = MethodCodeType.Runtime)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        void _VtblGap_27(); //390-363=27

        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1EF)> _
        //Sub RemoveDocumentInformation(<[In]()> ByVal RemoveDocInfoType As WdRemoveDocInfoType)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1F5)> _
        //Sub CheckInWithVersion(<[In]> ByVal Optional SaveChanges As Boolean = True, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Comments As Object, <[In]> ByVal Optional MakePublic As Boolean = False, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional VersionType As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1F9), TypeLibFunc(CShort(&H40))> _
        //Sub Dummy2()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), TypeLibFunc(CShort(&H40)), DispId(&H1FA)> _
        //Sub Dummy3()
        //<DispId(&H1FB)> _
        //ReadOnly Property ServerPolicy As <MarshalAs(UnmanagedType.Interface)> ServerPolicy
        //<DispId(&H1FC)> _
        //ReadOnly Property ContentControls As <MarshalAs(UnmanagedType.Interface)> ContentControls
        //<DispId(510)> _
        //ReadOnly Property DocumentInspectors As <MarshalAs(UnmanagedType.Interface)> DocumentInspectors
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1FD)> _
        //Sub LockServerFile()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1FF)> _
        //Function GetWorkflowTasks() As <MarshalAs(UnmanagedType.Interface)> WorkflowTasks
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H200)> _
        //Function GetWorkflowTemplates() As <MarshalAs(UnmanagedType.Interface)> WorkflowTemplates
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H202), TypeLibFunc(CShort(&H40))> _
        //Sub Dummy4()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H203), TypeLibFunc(CShort(&H40))> _
        //Sub AddMeetingWorkspaceHeader(<[In]()> ByVal SkipIfAbsent As Boolean, <[In](), MarshalAs(UnmanagedType.BStr)> ByVal Url As String, <[In](), MarshalAs(UnmanagedType.BStr)> ByVal Title As String, <[In](), MarshalAs(UnmanagedType.BStr)> ByVal Description As String, <[In](), MarshalAs(UnmanagedType.BStr)> ByVal ID As String)
        //<DispId(&H204)> _
        //ReadOnly Property Bibliography As <MarshalAs(UnmanagedType.Interface)> Bibliography
        //<DispId(&H205)> _
        //Property LockTheme As Boolean
        //<DispId(&H206)> _
        //Property LockQuickStyleSet As Boolean
        //<DispId(&H207)> _
        //ReadOnly Property OriginalDocumentTitle As <MarshalAs(UnmanagedType.BStr)> String
        //<DispId(520)> _
        //ReadOnly Property RevisedDocumentTitle As <MarshalAs(UnmanagedType.BStr)> String
        //<DispId(&H209)> _
        //ReadOnly Property CustomXMLParts As <MarshalAs(UnmanagedType.Interface)> CustomXMLParts
        //<DispId(&H20A)> _
        //Property FormattingShowNextLevel As Boolean
        //<DispId(&H20B)> _
        //Property FormattingShowUserStyleName As Boolean
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H20C)> _
        //Sub SaveAsQuickStyleSet(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal FileName As String)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H20D)> _
        //Sub ApplyQuickStyleSet(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal Name As String)
        //<DispId(&H20E)> _
        //ReadOnly Property Research As <MarshalAs(UnmanagedType.Interface)> Research
        //<DispId(&H20F)> _
        //Property Final As Boolean

        [PreserveSig]
        int Final_Get(out bool retVal); //390

        [PreserveSig]
        int Final_Let(bool sr);

        //<DispId(&H210)> _
        //Property OMathBreakBin As WdOMathBreakBin
        //<DispId(&H211)> _
        //Property OMathBreakSub As WdOMathBreakSub
        //<DispId(530)> _
        //Property OMathJc As WdOMathJc
        //<DispId(&H213)> _
        //Property OMathLeftMargin As Single
        //<DispId(&H214)> _
        //Property OMathRightMargin As Single
        //<DispId(&H217)> _
        //Property OMathWrap As Single
        //<DispId(&H218)> _
        //Property OMathIntSubSupLim As Boolean
        //<DispId(&H219)> _
        //Property OMathNarySupSubLim As Boolean
        //<DispId(&H21B)> _
        //Property OMathSmallFrac As Boolean
        //<DispId(&H21E)> _
        //ReadOnly Property WordOpenXML As <MarshalAs(UnmanagedType.BStr)> String
        //<DispId(&H221)> _
        //ReadOnly Property DocumentTheme As <MarshalAs(UnmanagedType.Interface)> OfficeTheme
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H222)> _
        //Sub ApplyDocumentTheme(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal FileName As String)
        //<DispId(&H224)> _
        //ReadOnly Property HasVBProject As Boolean
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H225)> _
        //Function SelectLinkedControls(<[In](), MarshalAs(UnmanagedType.Interface)> ByVal Node As CustomXMLNode) As <MarshalAs(UnmanagedType.Interface)> ContentControls
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(550)> _
        //Function SelectUnlinkedControls(<[In](), MarshalAs(UnmanagedType.Interface)> Optional ByVal Stream As CustomXMLPart = Nothing) As <MarshalAs(UnmanagedType.Interface)> ContentControls
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H227)> _
        //Function SelectContentControlsByTitle(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal Title As String) As <MarshalAs(UnmanagedType.Interface)> ContentControls
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H228)> _
        //Sub ExportAsFixedFormat(<[In], MarshalAs(UnmanagedType.BStr)> ByVal OutputFileName As String, <[In]> ByVal ExportFormat As WdExportFormat, <[In]> ByVal Optional OpenAfterExport As Boolean = False, <[In]> ByVal Optional OptimizeFor As WdExportOptimizeFor = 0, <[In]> ByVal Optional Range As WdExportRange = 0, <[In]> ByVal Optional From As Integer = 1, <[In]> ByVal Optional [To] As Integer = 1, <[In]> ByVal Optional Item As WdExportItem = 0, <[In]> ByVal Optional IncludeDocProps As Boolean = False, <[In]> ByVal Optional KeepIRM As Boolean = True, <[In]> ByVal Optional CreateBookmarks As WdExportCreateBookmarks = 0, <[In]> ByVal Optional DocStructureTags As Boolean = True, <[In]> ByVal Optional BitmapMissingFonts As Boolean = True, <[In]> ByVal Optional UseISO19005_1 As Boolean = False, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional FixedFormatExtClassPtr As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H229)> _
        //Sub FreezeLayout()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), TypeLibFunc(CShort(&H40)), DispId(&H22A)> _
        //Sub UnfreezeLayout()
        //<DispId(&H22B)> _
        //Property OMathFontName As <MarshalAs(UnmanagedType.BStr)> String
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H22E)> _
        //Sub DowngradeDocument()
        //<DispId(&H22F)> _
        //Property EncryptionProvider As <MarshalAs(UnmanagedType.BStr)> String
        //<DispId(560)> _
        //Property UseMathDefaults As Boolean
        //<DispId(&H233)> _
        //ReadOnly Property CurrentRsid As Integer
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H231)> _
        //Sub [Convert]()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H232)> _
        //Function SelectContentControlsByTag(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal Tag As String) As <MarshalAs(UnmanagedType.Interface)> ContentControls
    }
}