// © Copyright 2018 Levit & James, Inc.

using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace LevitJames.MSOffice.Internal
{
    [ComImport]
    [TypeLibType(flags: 0x10C0)]
    [Guid("000209A5-0000-0000-C000-000000000046")]
    internal interface View12 //_VtblGap7_194
    {
        //7 to get past the IDispatch members
        //The Activate member is member number 59 so that's 59-7 gives us 52 hence _VtblGap7_52 
        //Note: REMEMBER that read/write properties take 2 vtable slots.

        [MethodImpl(MethodCodeType = MethodCodeType.Runtime)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        void _VtblGap7_117();

        //<DispId(&H3E8)> _
        //ReadOnly Property Application As <MarshalAs(UnmanagedType.Interface)> Application
        //<DispId(&H3E9)> _
        //ReadOnly Property Creator As Integer
        //<DispId(&H3EA)> _
        //ReadOnly Property Parent As <MarshalAs(UnmanagedType.IDispatch)> Object
        //<DispId(0)> _
        //Default Property Type As WdViewType
        //<DispId(1)> _
        //Property FullScreen As Boolean
        //<DispId(2)> _
        //Property Draft As Boolean
        //<DispId(3)> _
        //Property ShowAll As Boolean
        //<DispId(4)> _
        //Property ShowFieldCodes As Boolean
        //<DispId(5)> _
        //Property MailMergeDataView As Boolean
        //<DispId(7)> _
        //Property Magnifier As Boolean
        //<DispId(8)> _
        //Property ShowFirstLineOnly As Boolean
        //<DispId(9)> _
        //Property ShowFormat As Boolean
        //<DispId(10)> _
        //ReadOnly Property Zoom As <MarshalAs(UnmanagedType.Interface)> Zoom
        //<DispId(11)> _
        //Property ShowObjectAnchors As Boolean
        //<DispId(12)> _
        //Property ShowTextBoundaries As Boolean
        //<DispId(13)> _
        //Property ShowHighlight As Boolean
        //<DispId(14)> _
        //Property ShowDrawings As Boolean
        //<DispId(15)> _
        //Property ShowTabs As Boolean
        //<DispId(&H10)> _
        //Property ShowSpaces As Boolean
        //<DispId(&H11)> _
        //Property ShowParagraphs As Boolean
        //<DispId(&H12)> _
        //Property ShowHyphens As Boolean
        //<DispId(&H13)> _
        //Property ShowHiddenText As Boolean
        //<DispId(20)> _
        //Property WrapToWindow As Boolean
        //<DispId(&H15)> _
        //Property ShowPicturePlaceHolders As Boolean
        //<DispId(&H16)> _
        //Property ShowBookmarks As Boolean
        //<DispId(&H17)> _
        //Property FieldShading As WdFieldShading
        //<DispId(&H18)> _
        //Property ShowAnimation As Boolean
        //<DispId(&H19)> _
        //Property TableGridlines As Boolean
        //<DispId(&H1A)> _
        //Property EnlargeFontsLessThan As Integer
        //<DispId(&H1B)> _
        //Property ShowMainTextLayer As Boolean
        //<DispId(&H1C)> _
        //Property SeekView As WdSeekView
        //<DispId(&H1D)> _
        //Property SplitSpecial As WdSpecialPane
        //<DispId(30)> _
        //Property BrowseToWindow As Integer
        //<DispId(&H1F)> _
        //Property ShowOptionalBreaks As Boolean
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H65)> _
        //Sub CollapseOutline(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Range As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H66)> _
        //Sub ExpandOutline(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Range As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H67)> _
        //Sub ShowAllHeadings()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H68)> _
        //Sub ShowHeading(<[In]()> ByVal Level As Integer)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H69)> _
        //Sub PreviousHeaderFooter()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H6A)> _
        //Sub NextHeaderFooter()
        //<DispId(&H20)> _
        //Property DisplayPageBoundaries As Boolean
        //<DispId(&H21)> _
        //Property DisplaySmartTags As Boolean
        //<DispId(&H22)> _
        //Property ShowRevisionsAndComments As Boolean
        //<DispId(&H23)> _
        //Property ShowComments As Boolean
        //<DispId(&H24)> _
        //Property ShowInsertionsAndDeletions As Boolean
        //<DispId(&H25)> _
        //Property ShowFormatChanges As Boolean
        //<DispId(&H26)> _
        //Property RevisionsView As WdRevisionsView
        //<DispId(&H27)> _
        //Property RevisionsMode As WdRevisionsMode
        //<DispId(40)> _
        //Property RevisionsBalloonWidth As Single
        //<DispId(&H29)> _
        //Property RevisionsBalloonWidthType As WdRevisionsBalloonWidthType
        //<DispId(&H2A)> _
        //Property RevisionsBalloonSide As WdRevisionsBalloonMargin
        //<DispId(&H2B)> _
        //ReadOnly Property Reviewers As <MarshalAs(UnmanagedType.Interface)> Reviewers
        //<DispId(&H2C)> _
        //Property RevisionsBalloonShowConnectingLines As Boolean
        //<DispId(&H2D)> _
        //Property ReadingLayout As Boolean
        //<DispId(&H2E)> _
        //Property ShowXMLMarkup As Integer
        //<DispId(&H2F)> _
        //Property ShadeEditableRanges As Integer
        //<DispId(&H30)> _
        //Property ShowInkAnnotations As Boolean
        //<DispId(&H31)> _
        //Property DisplayBackgrounds As Boolean
        //<DispId(50)> _
        //Property ReadingLayoutActualView As Boolean
        //<DispId(&H33)> _
        //Property ReadingLayoutAllowMultiplePages As Boolean
        //<DispId(&H35)> _
        //Property ReadingLayoutAllowEditing As Boolean
        //<DispId(&H36)> _
        //Property ReadingLayoutTruncateMargins As WdReadingLayoutMargin
        //<DispId(&H34)> _
        //Property ShowMarkupAreaHighlight As Boolean
        //<DispId(&H37)> _
        //Property Panning As Boolean

        // ''<DispId(&H38)> _
        // ''Property ShowCropMarks As Boolean
        [PreserveSig]
        int ShowCropMarks_Get(out bool retVal);

        [PreserveSig]
        int ShowCropMarks_Let(bool value);


        [DispId(dispId: 0x39)]
        WdRevisionsMode MarkupMode { get; set; }
    }
}