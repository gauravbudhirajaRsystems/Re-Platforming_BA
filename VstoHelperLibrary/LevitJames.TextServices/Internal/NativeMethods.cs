// © Copyright 2018 Levit & James, Inc.

using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

// ReSharper disable InconsistentNaming

namespace LevitJames.Interop
{
    [SuppressUnmanagedCodeSecurity]
    internal static partial class NativeMethods
    {
        private const string Kernel32 = "Kernel32";
        private const string User32 = "User32";
        private const string Shlwapi = "Shlwapi";


        public const int S_OK = 0x0;


        public const int ASSOCSTR_EXECUTABLE = 2;
        public const int ASSOCF_IGNOREUNKNOWN = 0x400;

        //Should return a Size_T (UIntLong) 
        //I don't expect hGlobals greater than an Int32 (Arrays cannot be resized using an Int64 anyway)
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage")]
        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return")]
        [DllImport(Kernel32, CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern int GlobalSize(IntPtr dataHandle);

        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage")]
        [DllImport(Kernel32, ExactSpelling = true)]
        public static extern IntPtr GlobalLock(SafeHGlobalHandle handle);

        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage")]
        [DllImport(Kernel32, ExactSpelling = true)]
        public static extern bool GlobalUnlock(SafeHGlobalHandle handle);

        [DllImport(Kernel32, ExactSpelling = true)]
        public static extern IntPtr GlobalFree(IntPtr handle);

        [DllImport(User32, CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int RegisterClipboardFormat(string format);

        [DllImport(Shlwapi, CharSet = CharSet.Unicode)]
        public static extern int AssocQueryString(int flags, int str, string pszAssoc, string pszExtra,
                                                  StringBuilder pszOut, ref int pcchOut);


        [DllImport(Kernel32, CharSet = CharSet.Unicode)]
        public static extern int GetModuleFileName(IntPtr hModule, StringBuilder lpFileName, int nSize);


        [DllImport(Kernel32, BestFitMapping = false, ExactSpelling = true, CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr LoadLibraryExW([In] string lpLibFileName, [In] IntPtr hFile, [In] uint dwFlags);


        [DllImport(Kernel32, CharSet = CharSet.Unicode)]
        public static extern IntPtr GetModuleHandle(string lpModuleName);
    }
}