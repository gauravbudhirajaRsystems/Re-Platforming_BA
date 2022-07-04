// © Copyright 2016 Levit & James, Inc.

using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security;
using System.Text;
using Microsoft.Win32.SafeHandles;

// ReSharper disable InconsistentNaming

namespace LevitJames.Core
{


    [SuppressMessage("Microsoft.Design", "CA1008:EnumsShouldHaveZeroValue")]
    [Flags]
    internal enum UserProfileTypes
    {
        Local = 0,// NativeMethods.PT_LOCAL,
        Temporary = 1,//NativeMethods.PT_TEMPORARY,
        Roaming = 2,//NativeMethods.PT_ROAMING,
        Mandatory = 3,//= NativeMethods.PT_MANDATORY
    }


    [SuppressUnmanagedCodeSecurity,
     SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
    internal static partial class NativeMethods
    {
        //Guard against dll's being loaded more than once due to case sensitivity
        //i.e user32 User32 User32.Dll are all valid but will load the same dll multiple times.
        private const string User32 = "user32";
        //private const string Gdi32 = "gdi32";
        private const string Kernel32 = "kernel32";
        private const string Ole32 = "ole32";
        private const string Mapi32 = "mapi32";
        private const string Shell32 = "shell32";
        private const string Shlwapi = "Shlwapi";
        private const string Wininet = "wininet";
        private const string Advapi32 = "advapi32";
        private const string Dnsapi = "Dnsapi";
        private const string UserEnv = "userenv";

        public const int MAPI_LOGON_UI = 0x1;
        public const int MAPI_DIALOG = 0x8;
        public const int MAX_INIFILE_ENTRY = 32768;
        public const int SW_MINIMIZE = 6;

        [DllImport(UserEnv)]
        public static extern bool GetProfileType(ref int pdwflags);


        [DllImport(Dnsapi, EntryPoint = "DnsQuery_W", ExactSpelling = true, CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern Int32 DnsQuery(string lpstrName, short wType, Int32 options, IntPtr pExtra, ref IntPtr ppQueryResultsSet, IntPtr pReserved);

        [DllImport(Dnsapi, SetLastError = true)]
        public static extern void DnsRecordListFree(IntPtr pRecordList, Int32 freeType);


        [DllImport(Advapi32, SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern int RegCopyTree(
            SafeRegistryHandle hKeySrc,
            [MarshalAs(UnmanagedType.LPWStr), Optional] string lpSubKey,
            SafeRegistryHandle hKeyDest
        );

 
        [DllImport(Advapi32, SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern uint RegSaveKey(SafeRegistryHandle hKey, string lpFile, IntPtr lpSecurityAttributes);


        // flags
        // REG_STANDARD_FORMAT  = 1
        // REG_LATEST_FORMAT    = 2;
        // REG_NO_COMPRESSION   = 3;

        [DllImport(Advapi32, SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern uint RegSaveKeyEx(SafeRegistryHandle hKey, string lpFile, IntPtr lpSecurityAttributes, int flags);


        #region      AssocQueryString 

        public const int ASSOCSTR_EXECUTABLE = 2;
        public const int ASSOCF_IGNOREUNKNOWN = 0x400;

        [DllImport(Shlwapi, CharSet = CharSet.Auto)]
        public static extern int AssocQueryString(int flags, int str, string pszAssoc, string pszExtra,
                                                  StringBuilder pszOut, ref int pcchOut);

        #endregion // AssocQueryString

        [DllImport(Wininet)]
        public static extern bool InternetGetConnectedState(out int description, int reservedValue);

        [DllImport(User32, CharSet = CharSet.Unicode)]
        public static extern int MessageBox(IntPtr h, string m, string c, int type);

        #region      LoadLibraryExW 

        [DllImport(Kernel32, BestFitMapping = false, ExactSpelling = true, CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr LoadLibraryExW([In] string lpLibFileName, [In] IntPtr hFile, [In] uint dwFlags);

        #endregion // LoadLibraryExW

        #region      GetProcAddress 

        [DllImport(Kernel32, CharSet = CharSet.Ansi, ExactSpelling = true)]
        public static extern IntPtr GetProcAddress(HandleRef hModule, string lpProcName);

        #endregion // GetProcAddress

        #region GetRunningObjectTable 

        [DllImport(Ole32)]
        public static extern void GetRunningObjectTable(int reserved,out IRunningObjectTable prot);

        #endregion // GetRunningObjectTable

        #region CreateBindCtx 

        [DllImport(Ole32)]
        public static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        #endregion // CreateBindCtx

        #region      GetKeyState 

        [DllImport(User32, SetLastError = false)]
        public static extern int GetKeyState(int vKey);

        #endregion // GetKeyState
        #region      GetKeyState 

        [DllImport(User32, SetLastError = false)]
        public static extern int GetAsyncKeyState(int vKey);

        #endregion // GetKeyState
        #region      MAPISendMail 

        [SuppressMessage("Microsoft.Globalization", "CA2101:SpecifyMarshalingForPInvokeStringArguments", MessageId = "MapiMessage.conversationID")]
        [SuppressMessage("Microsoft.Globalization", "CA2101:SpecifyMarshalingForPInvokeStringArguments", MessageId = "MapiMessage.dateReceived")]
        [SuppressMessage("Microsoft.Globalization", "CA2101:SpecifyMarshalingForPInvokeStringArguments", MessageId = "MapiMessage.messageType")]
        [SuppressMessage("Microsoft.Globalization", "CA2101:SpecifyMarshalingForPInvokeStringArguments", MessageId = "MapiMessage.noteText")]
        [SuppressMessage("Microsoft.Globalization", "CA2101:SpecifyMarshalingForPInvokeStringArguments", MessageId = "MapiMessage.subject")]
        [DllImport(Mapi32, CharSet = CharSet.Ansi, BestFitMapping = true, ThrowOnUnmappableChar = true)]
        public static extern int MAPISendMail(IntPtr session, IntPtr hwnd, [In] MapiMessage message, int flg, int rsv);

        #endregion // MAPISendMail

        #region      GetDesktopWindow 

        [DllImport(User32, SetLastError = false)]
        public static extern IntPtr GetDesktopWindow();

        #endregion // GetDesktopWindow

        #region      SHParseDisplayName 

        [DllImport(Shell32, EntryPoint = "SHParseDisplayName", ExactSpelling = true, BestFitMapping = false,
            ThrowOnUnmappableChar = true)]
        public static extern int SHParseDisplayName([MarshalAs(UnmanagedType.LPWStr)] string name, IntPtr bindingContext,
                                                    out IntPtr pidl, uint sfgaoIn, out uint sfgaoOut);

        #endregion // SHParseDisplayName

        #region      SHGetPathFromIDList 

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(Shell32, BestFitMapping = false, ThrowOnUnmappableChar = true)]
        public static extern bool SHGetPathFromIDList(IntPtr pidl,
                                                      [MarshalAs(UnmanagedType.LPTStr)] StringBuilder pszPath);

        #endregion // SHGetPathFromIDList

        #region      MapiMessage (class) 

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public class MapiMessage
        {
            public int reserved;
            [MarshalAs(UnmanagedType.LPStr)] public string subject;
            [MarshalAs(UnmanagedType.LPStr)] public string noteText;
            [MarshalAs(UnmanagedType.LPStr)] public string messageType;
            [MarshalAs(UnmanagedType.LPStr)] public string dateReceived;

            [MarshalAs(UnmanagedType.LPStr)] public string conversationID;
            public int flags;
            public IntPtr originator;
            public int recipCount;
            public IntPtr recips;
            public int fileCount;
            public IntPtr files;
        }

        #endregion // MapiMessage (class)

        #region      MapiFileDesc (class) 

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public class MapiFileDesc
        {
            public int reserved;
            public int flags;
            public int position;
            [MarshalAs(UnmanagedType.LPStr)] public string path;
            [MarshalAs(UnmanagedType.LPStr)] public string name;
            public IntPtr type;
        }

        #endregion // MapiFileDesc (class)

        #region      MapiRecipDesc (class) 

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public class MapiRecipDesc
        {
            public int reserved;
            public int recipClass;
            [MarshalAs(UnmanagedType.LPStr)] public string name;
            public IntPtr address;
            public int eIDSize;
            public IntPtr entryID;
        }

        #endregion // MapiRecipDesc (class)


        #region      GetPrivateProfileInt 

        [DllImport(Kernel32, CharSet=CharSet.Unicode, BestFitMapping=false, ThrowOnUnmappableChar=true)]
        public static extern int GetPrivateProfileInt(string lpApplicationName, string lpKeyName, int nDefault, string lpFileName);

        #endregion // GetPrivateProfileInt

        #region      GetPrivateProfileString 

        [DllImport(Kernel32, CharSet=CharSet.Unicode, BestFitMapping=false, ThrowOnUnmappableChar=true)]
        public static extern int GetPrivateProfileString(string lpApplicationName, string lpKeyName, string lpDefault, StringBuilder lpReturnedString, int nSize, string lpFileName);

        #endregion // GetPrivateProfileString

        #region      GetPrivateProfileSectionNames 

        [DllImport(Kernel32, CharSet=CharSet.Unicode, BestFitMapping=false, ThrowOnUnmappableChar=true)]
        public static extern int GetPrivateProfileSectionNames(byte[] lpszReturnBuffer, int nSize, string lpFileName);

        #endregion // GetPrivateProfileSectionNames

        #region      GetPrivateProfileSection 

        [DllImport(Kernel32, CharSet = CharSet.Unicode, BestFitMapping = false, ThrowOnUnmappableChar = true)]
        public static extern int GetPrivateProfileSection(string lpAppName, byte[] lpszReturnBuffer, int nSize, string lpFileName);

        #endregion // GetPrivateProfileSection

        [DllImport(User32)]
        public static extern int SetActiveWindow(IntPtr hWnd);

        [DllImport(User32)]
        public static extern IntPtr GetFocus();

        [DllImport(User32, CallingConvention = CallingConvention.StdCall)]
        public static extern IntPtr GetActiveWindow();

        [DllImport(User32, SetLastError = true)]
        public static extern IntPtr SetFocus(IntPtr hWnd);
      
        #region      ThrowWin32Error 

        //	Public Shared Sub ThrowWin32Error()
        //		ThrowWin32Exception(Nothing)
        //	End Sub

        //	<System.Security.SecurityCritical()>
        //<System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2201:DoNotRaiseReservedExceptionTypes")>
        //	Public Shared Sub ThrowWin32Exception(ByVal message As String)
        //		Dim er As int = Marshal.GetLastWin32Error()
        //		If er <> 0 Then
        //			If String.IsNullOrEmpty(message) Then
        //				Throw New System.ComponentModel.Win32Exception(er)
        //			Else
        //				Throw New System.ComponentModel.Win32Exception(er, message)
        //			End If
        //		End If
        //	End Sub

        #endregion // ThrowWin32Error
    }
}