// © Copyright 2018 Levit & James, Inc.

using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace LevitJames.MSOffice.Internal
{
    //<ComImport, Guid("0002096B-0000-0000-C000-000000000046"), CompilerGenerated, CoClass(GetType(Object)), TypeIdentifier> _

    [ComImport]
    [TypeLibType(flags: 0x1050)]
    [Guid("0002096B-0000-0000-C000-000000000046")]
    internal interface Document14
    {
        [DispId(dispId: 0)]
        string Name { get; }

        [MethodImpl(MethodCodeType = MethodCodeType.Runtime)]
        void _VtblGap1_425();

        [DispId(dispId: 0x237)]
        int CompatibilityMode { get; }
    }
}