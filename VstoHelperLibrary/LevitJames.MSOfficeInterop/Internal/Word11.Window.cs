// © Copyright 2018 Levit & James, Inc.

using System.ComponentModel;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace LevitJames.MSOffice.Internal
{
    [ComImport]
    [TypeLibType(flags: 0x10C0)]
    [Guid("00020962-0000-0000-C000-000000000046")]
    [DefaultMember("Caption")]
    internal interface Window11 //_VtblGap7_194
    {
        //7 to get past the IDispatch members
        //The Activate member is member number 59 so that's 59-7 gives us 52 hence _VtblGap7_52 
        //Note: REMEMBER that read/write properties take 2 vtable slots.

        [MethodImpl(MethodCodeType = MethodCodeType.Runtime)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        void _VtblGap7_52();

        //<DispId(&H3E8)> _
        //ReadOnly Property Application As <MarshalAs(UnmanagedType.Interface)> Application
        //<DispId(&H3E9)> _
        //ReadOnly Property Creator As Integer
        //<DispId(&H3EA)> _
        //ReadOnly Property Parent As <MarshalAs(UnmanagedType.IDispatch)> Object
        //<DispId(1)> _
        //ReadOnly Property ActivePane As <MarshalAs(UnmanagedType.Interface)> Pane
        //<DispId(2)> _
        //ReadOnly Property Document As <MarshalAs(UnmanagedType.Interface)> Document
        //<DispId(3)> _
        //ReadOnly Property Panes As <MarshalAs(UnmanagedType.Interface)> Panes
        //<DispId(4)> _
        //ReadOnly Property Selection As <MarshalAs(UnmanagedType.Interface)> Selection
        //<DispId(5)> _
        //Property Left As Integer
        //<DispId(6)> _
        //Property Top As Integer
        //<DispId(7)> _
        //Property Width As Integer
        //<DispId(8)> _
        //Property Height As Integer
        //<DispId(9)> _
        //Property Split As Boolean
        //<DispId(10)> _
        //Property SplitVertical As Integer
        //<DispId(0)> _
        //Property Caption As <MarshalAs(UnmanagedType.BStr)> String
        //<DispId(11)> _
        //Property WindowState As WdWindowState
        //<DispId(12)> _
        //Property DisplayRulers As Boolean
        //<DispId(13)> _
        //Property DisplayVerticalRuler As Boolean
        //<DispId(14)> _
        //ReadOnly Property View As <MarshalAs(UnmanagedType.Interface)> View
        //<DispId(15)> _
        //ReadOnly Property Type As WdWindowType
        //<DispId(&H10)> _
        //ReadOnly Property [Next] As <MarshalAs(UnmanagedType.Interface)> Window
        //<DispId(&H11)> _
        //ReadOnly Property Previous As <MarshalAs(UnmanagedType.Interface)> Window
        //<DispId(&H12)> _
        //ReadOnly Property WindowNumber As Integer
        //<DispId(&H13)> _
        //Property DisplayVerticalScrollBar As Boolean
        //<DispId(20)> _
        //Property DisplayHorizontalScrollBar As Boolean
        //<DispId(&H15)> _
        //Property StyleAreaWidth As Single
        //<DispId(&H16)> _
        //Property DisplayScreenTips As Boolean
        //<DispId(&H17)> _
        //Property HorizontalPercentScrolled As Integer
        //<DispId(&H18)> _
        //Property VerticalPercentScrolled As Integer
        //<DispId(&H19)> _
        //Property DocumentMap As Boolean
        //<DispId(&H1A)> _
        //ReadOnly Property Active As Boolean
        //<DispId(&H1B)> _
        //Property DocumentMapPercentWidth As Integer
        //<DispId(&H1C)> _
        //ReadOnly Property Index As Integer
        //<DispId(30)> _
        //Property IMEMode As WdIMEMode

        [PreserveSig]
        [DispId(dispId: 100)]
        int Activate();

        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H66)> _
        //Sub Close(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional SaveChanges As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional RouteDocument As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H67)> _
        //Sub LargeScroll(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Down As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Up As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ToRight As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ToLeft As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H68)> _
        //Sub SmallScroll(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Down As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Up As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ToRight As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ToLeft As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H69)> _
        //Function NewWindow() As <MarshalAs(UnmanagedType.Interface)> Window
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H6B), TypeLibFunc(CShort(&H40))> _
        //Sub PrintOutOld(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Background As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Append As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Range As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional OutputFileName As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional From As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional [To] As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Item As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Copies As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Pages As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PageType As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintToFile As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Collate As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ActivePrinterMacGX As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ManualDuplexPrint As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H6C)> _
        //Sub PageScroll(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Down As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Up As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H6D)> _
        //Sub SetFocus()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(110)> _
        //Function RangeFromPoint(<[In]()> ByVal x As Integer, <[In]()> ByVal y As Integer) As <MarshalAs(UnmanagedType.IDispatch)> Object
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H6F)> _
        //Sub ScrollIntoView(<[In], MarshalAs(UnmanagedType.IDispatch)> ByVal obj As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Start As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H70)> _
        //Sub GetPoint(<Out()> ByRef ScreenPixelsLeft As Integer, <Out()> ByRef ScreenPixelsTop As Integer, <Out()> ByRef ScreenPixelsWidth As Integer, <Out()> ByRef ScreenPixelsHeight As Integer, <[In](), MarshalAs(UnmanagedType.IDispatch)> ByVal obj As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), TypeLibFunc(CShort(&H40)), DispId(&H1BC)> _
        //Sub PrintOut2000(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Background As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Append As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Range As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional OutputFileName As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional From As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional [To] As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Item As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Copies As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Pages As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PageType As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintToFile As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Collate As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ActivePrinterMacGX As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ManualDuplexPrint As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomColumn As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomRow As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomPaperWidth As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomPaperHeight As Object)
        //<DispId(&H1F)> _
        //ReadOnly Property UsableWidth As Integer
        //<DispId(&H20)> _
        //ReadOnly Property UsableHeight As Integer
        //<DispId(&H21)> _
        //Property EnvelopeVisible As Boolean
        //<DispId(&H23)> _
        //Property DisplayRightRuler As Boolean
        //<DispId(&H22)> _
        //Property DisplayLeftScrollBar As Boolean
        //<DispId(&H24)> _
        //Property Visible As Boolean
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1BD)> _
        //Sub PrintOut(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Background As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Append As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Range As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional OutputFileName As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional From As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional [To] As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Item As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Copies As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Pages As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PageType As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintToFile As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Collate As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ActivePrinterMacGX As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ManualDuplexPrint As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomColumn As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomRow As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomPaperWidth As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PrintZoomPaperHeight As Object)
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H1BE)> _
        //Sub ToggleShowAllReviewers()
        //<DispId(&H25)> _
        //Property Thumbnails As Boolean
    }
}