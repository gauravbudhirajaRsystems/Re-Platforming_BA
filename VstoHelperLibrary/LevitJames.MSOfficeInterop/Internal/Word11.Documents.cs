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
    [TypeLibType(flags: 0x10C0)]
    [DefaultMember("Item")]
    [Guid("0002096C-0000-0000-C000-000000000046")]
    internal interface Documents
    {
        //Inherits IEnumerable
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(-4), TypeLibFunc(CShort(&H400))> _
        //Function GetEnumerator() As <MarshalAs(UnmanagedType.CustomMarshaler, MarshalType:="", MarshalTypeRef:=GetType(EnumeratorToEnumVariantMarshaler), MarshalCookie:="")> IEnumerator

        //7 to get past the IDispatch members and a count of 4 to reach the Result Item
        //Note: REMEMBER that read/write properties take 2 vtable slots.
        [MethodImpl(MethodCodeType = MethodCodeType.Runtime)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        void _VtblGap_5();

        //<DispId(2)> _
        //ReadOnly Property Count As Integer
        //<DispId(&H3E8)> _
        //ReadOnly Property Application As <MarshalAs(UnmanagedType.Interface)> Application
        //<DispId(&H3E9)> _
        //ReadOnly Property Creator As Integer
        //<DispId(&H3EA)> _
        //ReadOnly Property Parent As <MarshalAs(UnmanagedType.IDispatch)> Object
        [DispId(dispId: 0)]
        [PreserveSig]
        int Item(ref object Index, [MarshalAs(UnmanagedType.Interface)] ref Document rng);

        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H451)> _
        //  Sub Close(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional SaveChanges As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional OriginalFormat As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional RouteDocument As Object)
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), TypeLibFunc(CShort(&H40)), DispId(11)> _
        //  Function AddOld(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Template As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional NewTemplate As Object) As <MarshalAs(UnmanagedType.Interface)> Document
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(12), TypeLibFunc(CShort(&H40))> _
        //  Function OpenOld(<[In], MarshalAs(UnmanagedType.Struct)> ByRef FileName As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ConfirmConversions As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional [ReadOnly] As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional AddToRecentFiles As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PasswordDocument As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PasswordTemplate As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Revert As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional WritePasswordDocument As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional WritePasswordTemplate As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Format As Object) As <MarshalAs(UnmanagedType.Interface)> Document
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(13)> _
        //  Sub Save(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional NoPrompt As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional OriginalFormat As Object)
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(14)> _
        //  Function Add(<[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Template As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional NewTemplate As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional DocumentType As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Visible As Object) As <MarshalAs(UnmanagedType.Interface)> Document
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), TypeLibFunc(CShort(&H40)), DispId(15)> _
        //  Function Open2000(<[In], MarshalAs(UnmanagedType.Struct)> ByRef FileName As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ConfirmConversions As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional [ReadOnly] As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional AddToRecentFiles As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PasswordDocument As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PasswordTemplate As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Revert As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional WritePasswordDocument As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional WritePasswordTemplate As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Format As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Encoding As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Visible As Object) As <MarshalAs(UnmanagedType.Interface)> Document
        //		<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H10)> _
        //		Sub CheckOut(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal FileName As String)
        //		<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H11)> _
        //		Function CanCheckOut(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal FileName As String) As Boolean
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), TypeLibFunc(CShort(&H40)), DispId(&H12)> _
        //  Function Open2002(<[In], MarshalAs(UnmanagedType.Struct)> ByRef FileName As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ConfirmConversions As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional [ReadOnly] As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional AddToRecentFiles As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PasswordDocument As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PasswordTemplate As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Revert As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional WritePasswordDocument As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional WritePasswordTemplate As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Format As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Encoding As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Visible As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional OpenAndRepair As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional DocumentDirection As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional NoEncodingDialog As Object) As <MarshalAs(UnmanagedType.Interface)> Document
        //  <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(&H13)> _
        //  Function Open(<[In], MarshalAs(UnmanagedType.Struct)> ByRef FileName As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional ConfirmConversions As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional [ReadOnly] As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional AddToRecentFiles As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PasswordDocument As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional PasswordTemplate As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Revert As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional WritePasswordDocument As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional WritePasswordTemplate As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Format As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Encoding As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional Visible As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional OpenAndRepair As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional DocumentDirection As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional NoEncodingDialog As Object, <[In], MarshalAs(UnmanagedType.Struct)> ByRef Optional XMLTransform As Object) As <MarshalAs(UnmanagedType.Interface)> Document
    }
}