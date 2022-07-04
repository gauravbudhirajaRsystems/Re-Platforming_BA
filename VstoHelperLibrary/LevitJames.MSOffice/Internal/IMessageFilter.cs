// © Copyright 2018 Levit & James, Inc.

using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace LevitJames.Libraries
{
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    internal struct INTERFACEINFO
    {
        [MarshalAs(UnmanagedType.IUnknown)] public object punk;
        public Guid iid;
        public ushort wMethod;
    }

    [ComImport]
    [ComConversionLoss]
    [InterfaceType(interfaceType: 1)]
    [Guid("00000016-0000-0000-C000-000000000046")]
    internal interface IMessageFilter
    {
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        int HandleInComingCall(uint dwCallType, IntPtr htaskCaller, uint dwTickCount,
                               [MarshalAs(UnmanagedType.LPArray)] INTERFACEINFO[] lpInterfaceInfo);

        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        int RetryRejectedCall(IntPtr htaskCallee, uint dwTickCount, uint dwRejectType);

        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        int MessagePending(IntPtr htaskCallee, uint dwTickCount, uint dwPendingType);
    }
}