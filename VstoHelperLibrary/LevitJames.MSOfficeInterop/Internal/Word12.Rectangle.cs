// © Copyright 2018 Levit & James, Inc.

using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace LevitJames.MSOffice.Internal
{
    [ComImport]
    [TypeLibType(flags: 0x10C0)]
    [Guid("ADD4EDF3-2F33-4734-9CE6-D476097C5ADA")]
    internal interface Rectangle12 //_VtblGap7_194
    {
        [MethodImpl(MethodCodeType = MethodCodeType.Runtime)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        void _VtblGap7_3();

        //<DispId(&H3E8)> _
        //ReadOnly Property Application As <MarshalAs(UnmanagedType.Interface)> Application
        //<DispId(&H3E9)> _
        //ReadOnly Property Creator As Integer
        //<DispId(&H3EA)> _
        //ReadOnly Property Parent As <MarshalAs(UnmanagedType.IDispatch)> Object

        [DispId(dispId: 0x2)]
        [PreserveSig]
        int RectangleType(out WdRectangleType rt);

        [MethodImpl(MethodCodeType = MethodCodeType.Runtime)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        void _VtblGap_4();

        //<DispId(3)> _
        //ReadOnly Property Left As Integer
        //<DispId(4)> _
        //ReadOnly Property Top As Integer
        //<DispId(5)> _
        //ReadOnly Property Width As Integer
        //<DispId(6)> _
        //ReadOnly Property Height As Integer

        [DispId(dispId: 0x7)]
        [PreserveSig]
        int Range([Out] [MarshalAs(UnmanagedType.Interface)]
                  out Range sr);

        //<DispId(8)> _
        //ReadOnly Property Lines As <MarshalAs(UnmanagedType.Interface)> Lines
    }
}