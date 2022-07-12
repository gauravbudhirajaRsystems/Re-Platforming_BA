using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace LevitJames.MSOfficeInterop.Common.Internal
{
    [ComImport]
    [TypeLibType(flags: 0x10C0)]
    [Guid("ADD4EDF3-2F33-4734-9CE6-D476097C5ADA")]
    internal interface Rectangle12 //_VtblGap7_194
    {
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
    }
}
