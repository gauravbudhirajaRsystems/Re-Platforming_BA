// © Copyright 2018 Levit & James, Inc.

using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace LevitJames.MSOffice.Internal
{
    [ComImport]
    [Guid("0002095E-0000-0000-C000-000000000046")]
    [TypeLibType(flags: 0x10C0)]
    internal interface Range11
    {
        //7 to get past the IDispatch members

        //The StoryType member is member number 18 so that's 18-7 gives us 11 hence _VtblGap7_73 
        //Note: REMEMBER that read/write properties take 2 vtable slots.
        [MethodImpl(MethodCodeType = MethodCodeType.Runtime)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        void _VtblGap7_11();

        [DispId(dispId: 0x1A8)]
        [PreserveSig]
        int StoryType([Out] out WdStoryType st);

        //The ShapeRange member is member number 80 so that's (80-(18+1)) gives us 61 hence _VtblGap_61 
        //Note: REMEMBER that read/write properties take 2 vtable slots.
        [MethodImpl(MethodCodeType = MethodCodeType.Runtime)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        void _VtblGap_61();

        [DispId(dispId: 0x1A8)]
        [PreserveSig]
        int ShapeRange([Out] [MarshalAs(UnmanagedType.Interface)]
                       out ShapeRange sr);

        [DispId(dispId: 0x138)]
        [PreserveSig]
        int Case_Get(out int retval);
    }
}