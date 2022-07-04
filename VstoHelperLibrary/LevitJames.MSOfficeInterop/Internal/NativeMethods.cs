// © Copyright 2018 Levit & James, Inc.

using System;
using System.CodeDom.Compiler;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

// ReSharper disable InconsistentNaming
namespace LevitJames.MSOffice.Internal
{
    [GeneratedCode("", "")]
    [SuppressUnmanagedCodeSecurity]
    internal static class NativeMethods
    {
        //Guard against dll's being loaded more than once due to case sensitivity
        //i.e user32 User32 User32.Dll are all valid but will load the same dll multiple times.
        //private const string User32 = "user32";
        //private const string Gdi32 = "gdi32";
        private const string Kernel32 = "kernel32";
        //private const string UXtheme = "uxtheme";
        //private const string ComCtl32 = "comctl32";
        //private const string MsImg32 = "msimg32";
        //private const string OleAcc = "oleacc";
        //private const string Shell32 = "shell32";
        //private const string Shlwapi = "shlwapi";
        //private const string UserEnv = "userenv";
        //private const string Ole32 = "ole32";

        public const string SystemDialogClassName = "#32770";

 

        [DllImport(Kernel32, CharSet = CharSet.Auto)]
        public static extern IntPtr GetModuleHandle(string lpModuleName);


        [DllImport(Kernel32, CharSet = CharSet.Unicode)]
        public static extern int GetModuleFileName(IntPtr hModule, StringBuilder lpFileName, int nSize);

        [StructLayout(LayoutKind.Sequential)]
        public struct RGBQUAD
        {
            public byte rgbBlue;
            public byte rgbGreen;
            public byte rgbRed;
            public byte rgbReserved;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct BITMAP
        {
            public int bmType;
            public int bmWidth;
            public int bmHeight;
            public int bmWidthBytes;
            public short bmPlanes;
            public short bmBitsPixel;
            public IntPtr bmBits;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct BITMAPINFOHEADER
        {
            public int biSize;
            public int biWidth;
            public int biHeight;
            public short biPlanes;
            public short biBitCount;
            public int biCompression;
            public int biSizeImage;
            public int biXPelsPerMeter;
            public int biYPelsPerMeter;
            public int biClrUsed;
            public int bitClrImportant;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct DIBSECTION
        {
            public BITMAP dsBm;
            public BITMAPINFOHEADER dsBmih;
            public int dsBitField1;
            public int dsBitField2;
            public int dsBitField3;
            public IntPtr dshSection;
            public int dsOffset;
        }

        [DllImport("gdi32.dll", EntryPoint = "GetObject")]
        public static extern int GetObjectDIBSection(IntPtr hObject, int nCount, ref DIBSECTION lpObject);

        [DllImport("gdi32.dll")]
        public static extern void DeleteObject(IntPtr handle);


    }
}