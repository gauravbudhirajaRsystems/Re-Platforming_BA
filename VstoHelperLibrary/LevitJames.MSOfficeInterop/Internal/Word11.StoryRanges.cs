// © Copyright 2018 Levit & James, Inc.

using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

// ReSharper disable InconsistentNaming

namespace LevitJames.MSOffice.Internal
{
    [ComImport]
    [TypeLibType(flags: 0x10C0)]
    [Guid("0002098C-0000-0000-C000-000000000046")]
    internal interface StoryRanges11
    {
        //Inherits IEnumerable
        //<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), DispId(-4), TypeLibFunc(CShort(&H400))> 
        //Function GetEnumerator() As <MarshalAs(UnmanagedType.CustomMarshaler, MarshalType:="", MarshalTypeRef:=GetType(EnumeratorToEnumVariantMarshaler), MarshalCookie:="")> IEnumerator

        //7 to get past the IDispatch members and a count of 4 to reach the Result Item
        //Note: REMEMBER that read/write properties take 2 vtable slots.
        [MethodImpl(MethodCodeType = MethodCodeType.Runtime)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        void _VtblGap_5();

        //<DispId(2)> 
        //ReadOnly Property Count() As Integer
        //<DispId(&H3E8)> 
        //ReadOnly Property Application() As <MarshalAs(UnmanagedType.Interface)> Application
        //<DispId(&H3E9)> 
        //ReadOnly Property Creator() As Integer
        //<DispId(&H3EA)> 
        //ReadOnly Property Parent() As <MarshalAs(UnmanagedType.IDispatch)> Object
        [DispId(dispId: 0)]
        [PreserveSig]
        int Item(WdStoryType Index, [MarshalAs(UnmanagedType.Interface)] out Range rng);
    }
}