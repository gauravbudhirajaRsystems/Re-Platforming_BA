// © Copyright 2018 Levit & James, Inc.

using System.ComponentModel;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

// ReSharper disable InconsistentNaming

namespace LevitJames.MSOffice.Internal
{
    [ComImport]
    [Guid("0002092D-0000-0000-C000-000000000046")]
    [DefaultMember("Item")]
    [TypeLibType(flags: 0x10C0)]
    internal interface Styles11
    {
        //Inherits IEnumerable
        //<DispId(&H3E8)> _
        //ReadOnly Property Application As <MarshalAs(UnmanagedType.Interface)> Application
        //<DispId(&H3E9)> _
        //ReadOnly Property Creator As Integer
        //<DispId(&H3EA)> _
        //ReadOnly Property Parent As <MarshalAs(UnmanagedType.IDispatch)> Object
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), TypeLibFunc(CShort(&H400)), DispId(-4)> _
        //Function GetEnumerator() As <MarshalAs(UnmanagedType.CustomMarshaler, MarshalType:="", MarshalTypeRef:=GetType(EnumeratorToEnumVariantMarshaler), MarshalCookie:="")> IEnumerator
        //<DispId(1)> _
        //ReadOnly Property Count As Integer
        //<DispId(0)> _

        //Inherits IEnumerable
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(-4), TypeLibFunc(CShort(&H400))> 
        //Function GetEnumerator() As <MarshalAs(UnmanagedType.CustomMarshaler, MarshalType:="", MarshalTypeRef:=GetType(EnumeratorToEnumVariantMarshaler), MarshalCookie:="")> IEnumerator

        //7 to get past the IDispatch members and a count of 4 to reach the Result Item
        //Note: REMEMBER that read/write properties take 2 vtable slots.
        [MethodImpl(MethodCodeType = MethodCodeType.Runtime)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        void _VtblGap_5();

        [PreserveSig]
        int Item(ref object Index, [MarshalAs(UnmanagedType.Interface)] ref Style style);

        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(100)> _
        //Function Add(<[In], MarshalAs(UnmanagedType.BStr)> ByVal Name As String, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Type As Object) As <MarshalAs(UnmanagedType.Interface)> Style
    }
}