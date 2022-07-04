// © Copyright 2018 Levit & James, Inc.

using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace LevitJames.MSOffice.Internal
{
    [ComImport]
    [TypeLibType(flags: 0x10C0)]
    [Guid("0002092F-0000-0000-C000-000000000046")]
    internal interface Field11
    {
        //7 to get past the IDispatch members and a count of 7 to reach the Result Member
        //Result is at Offset 16 -7 to get past the IDispatch=  9
        //Note: REMEMBER that read/write properties take 2 vtable slots.
        [MethodImpl(MethodCodeType = MethodCodeType.Runtime)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        void _VtblGap_9();

        // <DispId(&H3E8)> 
        // ReadOnly Property Application() As <MarshalAs(UnmanagedType.Interface)> Application
        // <DispId(&H3E9)> 
        // ReadOnly Property Creator() As Integer
        // <DispId(&H3EA)> 
        // ReadOnly Property Parent() As <MarshalAs(UnmanagedType.IDispatch)> Object
        // <DispId(0)> 
        //Property Code() As <MarshalAs(UnmanagedType.Interface)> Range
        // <DispId(1)> 
        // ReadOnly Property Type() As WdFieldType
        // <DispId(2)> 
        // Property Locked() As Boolean
        // <DispId(3)> 
        // ReadOnly Property Kind() As WdFieldKind
        [DispId(dispId: 4)]
        [PreserveSig]
        int Result([MarshalAs(UnmanagedType.Interface)] ref Range retVal);

        //<DispId(4), PreserveSig()> 
        //Function Result(<MarshalAs(UnmanagedType.Interface)> ByRef retVal As Range) As int
        //<DispId(5)> 
        //Property Data() As <MarshalAs(UnmanagedType.BStr)> String
        //<DispId(6)> 
        //ReadOnly Property [Next]() As <MarshalAs(UnmanagedType.Interface)> Field
        //<DispId(7)> 
        //ReadOnly Property Previous() As <MarshalAs(UnmanagedType.Interface)> Field
        //<DispId(8)> 
        //ReadOnly Property Index() As Integer
        //<DispId(9)> 
        //Property ShowCodes() As Boolean
        //<DispId(10)> 
        //ReadOnly Property LinkFormat() As <MarshalAs(UnmanagedType.Interface)> LinkFormat
        //<DispId(11)> 
        //ReadOnly Property OLEFormat() As <MarshalAs(UnmanagedType.Interface)> OLEFormat
        //<DispId(12)> 
        //ReadOnly Property InlineShape() As <MarshalAs(UnmanagedType.Interface)> InlineShape
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&HFFFF)> 
        //Sub [Select]()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H65)> 
        //Function Update() As Boolean
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H66)> 
        //Sub Unlink()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H67)> 
        //Sub UpdateSource()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H68)> 
        //Sub DoClick()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H69)> 
        //Sub Copy()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H6A)> 
        //Sub Cut()
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H6B)> 
        //Sub Delete()
    }
}