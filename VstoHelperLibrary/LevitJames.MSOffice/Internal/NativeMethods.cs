// © Copyright 2018 Levit & James, Inc.

using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Permissions;
using System.Text;
using LevitJames.Core;
using Microsoft.Office.Core;
// ReSharper disable UnusedMember.Local

// ReSharper disable InconsistentNaming
namespace LevitJames.Libraries
{
    [GeneratedCode("", "")]
    [SuppressUnmanagedCodeSecurity]
    internal static class NativeMethods
    {
        public delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);


        public delegate bool EnumThreadWndProc(IntPtr hWnd, IntPtr lParam);


        //[GeneratedCode("", "")]
        //[CompilerGenerated]
        //public struct WINDOWPLACEMENT
        //{
        //    public int length;
        //    public int flags;
        //    public int showCmd;
        //    public Point ptMinPosition;
        //    public Point ptMaxPosition;
        //    public Rectangle rcNormalPosition;
        //}

        //[return: MarshalAs(UnmanagedType.Bool)]
        //[DllImport(User32)]
        //public static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);


        public delegate IntPtr fnHookProc(int nCode, IntPtr wParam, IntPtr lParam);


        public enum DPI_AWARENESS_CONTEXT
        {
            DPI_AWARENESS_CONTEXT_UNAWARE = -1,
            DPI_AWARENESS_CONTEXT_SYSTEM_AWARE = -2,
            DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE = -3,
            DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
        }


        public enum DrawFrameControlState
        {
            //Button
            Button3DState = 0x8,
            ButtonCheck = 0x0,
            ButtonPush = 0x10,

            ButtonRadioImage = 0x1,
            ButtonRadioMask = 0x2,
            ButtonRadio = 0x4,

            //Caption
            CaptionClose = 0x0,
            CaptionMin = 0x1,
            CaptionMax = 0x2,
            CaptionRestore = 0x3,
            CaptionHelp = 0x4,

            //Menu
            MenuArrow = 0x0,
            MenuBullet = 0x2,
            MenuArrowDown = 0x10,
            MenuArrowUp = 0x8,

            //ScrollBar
            ScrollUp = 0x0,
            ScrollDown = 0x1,
            ScrollLeft = 0x2,
            ScrollRight = 0x3,
            ScrollComboBox = 0x5,

            ScrollSizeGrip = 0x8,
            ScrollSizeGripRight = 0x10,

            //States
            Inactive = 0x100,
            Pushed = 0x200,
            Checked = 0x400,
            Transparent = 0x800,
            Hot = 0x1000,
            AdjustRect = 0x2000,
            Flat = 0x4000,
            Mono = 0x8000
        }


        public enum DrawFrameControlType
        {
            Caption = 1,
            Menu = 2,
            Scroll = 3,
            Button = 4,
            PopupMenu = 5
        }


        [Flags]
        public enum DrawStyleFlags
        {
            ILD_NORMAL = 0x0,
            ILD_TRANSPARENT = 0x1,
            ILD_MASK = 0x10,
            ILD_IMAGE = 0x20,

            //#If (WIN32_IE >= &H300) Then
            ILD_ROP = 0x40,

            //#End If
            ILD_BLEND25 = 0x2,
            ILD_BLEND50 = 0x4,
            ILD_OVERLAYMASK = 0xF00,

            ILD_SELECTED = ILD_BLEND50,
            ILD_FOCUS = ILD_BLEND25,
            ILD_BLEND = ILD_BLEND50
        }


        [Flags]
        public enum ExtTextOutOptions
        {
            ETO_GRAYED = 0x1,
            ETO_OPAQUE = 0x2,
            ETO_CLIPPED = 0x4,
            ETO_GLYPH_INDEX = 0x10,
            ETO_RTLREADING = 0x80
        }


        [Flags]
        public enum GAFlags
        {
            Parent = 1,
            Root = 2,
            RootOwner = 3
        }


        public enum GetWindowCommand
        {
            GW_CHILD = 5,
            GW_HWNDFIRST = 0,
            GW_HWNDLAST = 1,
            GW_HWNDNEXT = 2,
            GW_HWNDPREV = 3,
            GW_OWNER = 4
        }


        [Flags]
        public enum GUITHREADINFOFlags
        {
            GUI_CARETBLINKING = 0x1,
            GUI_INMOVESIZE = 0x2,
            GUI_SYSTEMMENUMODE = 0x8,
            GUI_INMENUMODE = 0x4,
            GUI_POPUPMENUMODE = 0x10,
            GUI_16BITTASK = 0x20
        }


        public enum HitTestResult
        {
            HTBORDER = 18,
            HTBOTTOM = 15,
            HTBOTTOMLEFT = 16,
            HTBOTTOMRIGHT = 17,
            HTCAPTION = 2,
            HTCLIENT = 1,
            HTCLOSE = 20,
            HTERROR = -2,
            HTGROWBOX = 4,
            HTHELP = 21,
            HTHSCROLL = 6,
            HTLEFT = 10,
            HTMAXBUTTON = 9,
            HTMENU = 5,
            HTMINBUTTON = 8,
            HTNOWHERE = 0,
            HTOBJECT = 19,
            HTREDUCE = HTMINBUTTON,
            HTRIGHT = 11,
            HTSIZE = HTGROWBOX,
            HTSIZEFIRST = HTLEFT,
            HTSIZELAST = HTBOTTOMRIGHT,
            HTSYSMENU = 3,
            HTTOP = 12,
            HTTOPLEFT = 13,
            HTTOPRIGHT = 14,
            HTTRANSPARENT = -1,
            HTVSCROLL = 7,
            HTZOOM = HTMAXBUTTON
        }


        [Flags]
        public enum RedrawWindowFlags
        {
            RDW_ALLCHILDREN = 0x80,
            RDW_ERASE = 0x4,
            RDW_ERASENOW = 0x200,
            RDW_FRAME = 0x400,
            RDW_INTERNALPAINT = 0x2,
            RDW_INVALIDATE = 0x1,
            RDW_NOCHILDREN = 0x40,
            RDW_NOERASE = 0x20,
            RDW_NOFRAME = 0x800,
            RDW_NOINTERNALPAINT = 0x10,
            RDW_UPDATENOW = 0x100,
            RDW_VALIDATE = 0x8
        }


        [Flags]
        public enum SetWindowPosFlags
        {
            SWP_ASYNCWINDOWPOS = 0x4000,
            SWP_DEFERERASE = 0x2000,
            SWP_DRAWFRAME = SWP_FRAMECHANGED,
            SWP_FRAMECHANGED = 0x20,
            SWP_HIDEWINDOW = 0x80,
            SWP_NOACTIVATE = 0x10,
            SWP_NOCOPYBITS = 0x100,
            SWP_NOMOVE = 0x2,
            SWP_NOOWNERZORDER = 0x200,
            SWP_NOREDRAW = 0x8,
            SWP_NOREPOSITION = SWP_NOOWNERZORDER,
            SWP_NOSIZE = 0x1,
            SWP_NOZORDER = 0x4,
            SWP_SHOWWINDOW = 0x40,
            SWP_NOSIZE_NOMOVE = SWP_NOSIZE | SWP_NOMOVE,
            SWP_NOSIZE_NOMOVE_NOACTIVATE = SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE
        }


        public enum ShowWindowFlags
        {
            SW_HIDE = 0,
            SW_SHOWNORMAL = 1,
            SW_SHOWMINIMIZED = 2,
            SW_MAXIMIZE = 3,
            SW_SHOWMAXIMIZED = 3,
            SW_SHOWNOACTIVATE = 4,
            SW_SHOW = 5,
            SW_MINIMIZE = 6,
            SW_SHOWMINNOACTIVE = 7,
            SW_SHOWNA = 8,
            SW_RESTORE = 9,
            SW_SHOWDEFAULT = 10
        }


        public enum SystemParametersInfoEnum
        {
            SPI_GETBEEP = 0x1,
            SPI_SETBEEP = 0x2,
            SPI_GETMOUSE = 0x3,
            SPI_SETMOUSE = 0x4,
            SPI_GETBORDER = 0x5,
            SPI_SETBORDER = 0x6,
            SPI_GETKEYBOARDSPEED = 0xA,
            SPI_SETKEYBOARDSPEED = 0xB,
            SPI_LANGDRIVER = 0xC,
            SPI_ICONHORIZONTALSPACING = 0xD,
            SPI_GETSCREENSAVETIMEOUT = 0xE,
            SPI_SETSCREENSAVETIMEOUT = 0xF,
            SPI_GETSCREENSAVEACTIVE = 0x10,
            SPI_SETSCREENSAVEACTIVE = 0x11,
            SPI_GETGRIDGRANULARITY = 0x12,
            SPI_SETGRIDGRANULARITY = 0x13,
            SPI_SETDESKWALLPAPER = 0x14,
            SPI_SETDESKPATTERN = 0x15,
            SPI_GETKEYBOARDDELAY = 0x16,
            SPI_SETKEYBOARDDELAY = 0x17,
            SPI_ICONVERTICALSPACING = 0x18,
            SPI_GETICONTITLEWRAP = 0x19,
            SPI_SETICONTITLEWRAP = 0x1A,
            SPI_GETMENUDROPALIGNMENT = 0x1B,
            SPI_SETMENUDROPALIGNMENT = 0x1C,
            SPI_SETDOUBLECLKWIDTH = 0x1D,
            SPI_SETDOUBLECLKHEIGHT = 0x1E,
            SPI_GETICONTITLELOGFONT = 0x1F,
            SPI_SETDOUBLECLICKTIME = 0x20,
            SPI_SETMOUSEBUTTONSWAP = 0x21,
            SPI_SETICONTITLELOGFONT = 0x22,
            SPI_GETFASTTASKSWITCH = 0x23,
            SPI_SETFASTTASKSWITCH = 0x24,

            //#if(WINVER >= =&H0400)
            SPI_SETDRAGFULLWINDOWS = 0x25,
            SPI_GETDRAGFULLWINDOWS = 0x26,
            SPI_GETNONCLIENTMETRICS = 0x29,
            SPI_SETNONCLIENTMETRICS = 0x2A,
            SPI_GETMINIMIZEDMETRICS = 0x2B,
            SPI_SETMINIMIZEDMETRICS = 0x2C,
            SPI_GETICONMETRICS = 0x2D,
            SPI_SETICONMETRICS = 0x2E,
            SPI_SETWORKAREA = 0x2F,
            SPI_GETWORKAREA = 0x30,
            SPI_SETPENWINDOWS = 0x31,

            SPI_GETHIGHCONTRAST = 0x42,
            SPI_SETHIGHCONTRAST = 0x43,
            SPI_GETKEYBOARDPREF = 0x44,
            SPI_SETKEYBOARDPREF = 0x45,
            SPI_GETSCREENREADER = 0x46,
            SPI_SETSCREENREADER = 0x47,
            SPI_GETANIMATION = 0x48,
            SPI_SETANIMATION = 0x49,
            SPI_GETFONTSMOOTHING = 0x4A,
            SPI_SETFONTSMOOTHING = 0x4B,
            SPI_SETDRAGWIDTH = 0x4C,
            SPI_SETDRAGHEIGHT = 0x4D,
            SPI_SETHANDHELD = 0x4E,
            SPI_GETLOWPOWERTIMEOUT = 0x4F,
            SPI_GETPOWEROFFTIMEOUT = 0x50,
            SPI_SETLOWPOWERTIMEOUT = 0x51,
            SPI_SETPOWEROFFTIMEOUT = 0x52,
            SPI_GETLOWPOWERACTIVE = 0x53,
            SPI_GETPOWEROFFACTIVE = 0x54,
            SPI_SETLOWPOWERACTIVE = 0x55,
            SPI_SETPOWEROFFACTIVE = 0x56,
            SPI_SETCURSORS = 0x57,
            SPI_SETICONS = 0x58,
            SPI_GETDEFAULTINPUTLANG = 0x59,
            SPI_SETDEFAULTINPUTLANG = 0x5A,
            SPI_SETLANGTOGGLE = 0x5B,
            SPI_GETWINDOWSEXTENSION = 0x5C,
            SPI_SETMOUSETRAILS = 0x5D,
            SPI_GETMOUSETRAILS = 0x5E,
            SPI_SETSCREENSAVERRUNNING = 0x61,
            SPI_SCREENSAVERRUNNING = SPI_SETSCREENSAVERRUNNING,

            //#endif /* WINVER >= =&H0400 */
            SPI_GETFILTERKEYS = 0x32,
            SPI_SETFILTERKEYS = 0x33,
            SPI_GETTOGGLEKEYS = 0x34,
            SPI_SETTOGGLEKEYS = 0x35,
            SPI_GETMOUSEKEYS = 0x36,
            SPI_SETMOUSEKEYS = 0x37,
            SPI_GETSHOWSOUNDS = 0x38,
            SPI_SETSHOWSOUNDS = 0x39,
            SPI_GETSTICKYKEYS = 0x3A,
            SPI_SETSTICKYKEYS = 0x3B,
            SPI_GETACCESSTIMEOUT = 0x3C,
            SPI_SETACCESSTIMEOUT = 0x3D,

            //#if(WINVER >= =&H0400)
            SPI_GETSERIALKEYS = 0x3E,
            SPI_SETSERIALKEYS = 0x3F,

            //#endif /* WINVER >= =&H0400 */
            SPI_GETSOUNDSENTRY = 0x40,
            SPI_SETSOUNDSENTRY = 0x41,

            //#if(_WIN32_WINNT >= =&H0400)
            SPI_GETSNAPTODEFBUTTON = 0x5F,
            SPI_SETSNAPTODEFBUTTON = 0x60,

            //#endif /* _WIN32_WINNT >= =&H0400 */
            //#if (_WIN32_WINNT >= =&H0400) || (_WIN32_WINDOWS > =&H0400)
            SPI_GETMOUSEHOVERWIDTH = 0x62,
            SPI_SETMOUSEHOVERWIDTH = 0x63,
            SPI_GETMOUSEHOVERHEIGHT = 0x64,
            SPI_SETMOUSEHOVERHEIGHT = 0x65,
            SPI_GETMOUSEHOVERTIME = 0x66,
            SPI_SETMOUSEHOVERTIME = 0x67,
            SPI_GETWHEELSCROLLLINES = 0x68,
            SPI_SETWHEELSCROLLLINES = 0x69,
            SPI_GETMENUSHOWDELAY = 0x6A,
            SPI_SETMENUSHOWDELAY = 0x6B,


            SPI_GETSHOWIMEUI = 0x6E,
            SPI_SETSHOWIMEUI = 0x6F,
            //#End If


            //#if(WINVER >= =&H0500)
            SPI_GETMOUSESPEED = 0x70,
            SPI_SETMOUSESPEED = 0x71,
            SPI_GETSCREENSAVERRUNNING = 0x72,
            SPI_GETDESKWALLPAPER = 0x73,
            //#endif /* WINVER >= =&H0500 */


            //#if(WINVER >= =&H0500)
            SPI_GETACTIVEWINDOWTRACKING = 0x1000,
            SPI_SETACTIVEWINDOWTRACKING = 0x1001,
            SPI_GETMENUANIMATION = 0x1002,
            SPI_SETMENUANIMATION = 0x1003,
            SPI_GETCOMBOBOXANIMATION = 0x1004,
            SPI_SETCOMBOBOXANIMATION = 0x1005,
            SPI_GETLISTBOXSMOOTHSCROLLING = 0x1006,
            SPI_SETLISTBOXSMOOTHSCROLLING = 0x1007,
            SPI_GETGRADIENTCAPTIONS = 0x1008,
            SPI_SETGRADIENTCAPTIONS = 0x1009,
            SPI_GETKEYBOARDCUES = 0x100A,
            SPI_SETKEYBOARDCUES = 0x100B,
            SPI_GETMENUUNDERLINES = SPI_GETKEYBOARDCUES,
            SPI_SETMENUUNDERLINES = SPI_SETKEYBOARDCUES,
            SPI_GETACTIVEWNDTRKZORDER = 0x100C,
            SPI_SETACTIVEWNDTRKZORDER = 0x100D,
            SPI_GETHOTTRACKING = 0x100E,
            SPI_SETHOTTRACKING = 0x100F,
            SPI_GETMENUFADE = 0x1012,
            SPI_SETMENUFADE = 0x1013,
            SPI_GETSELECTIONFADE = 0x1014,
            SPI_SETSELECTIONFADE = 0x1015,
            SPI_GETTOOLTIPANIMATION = 0x1016,
            SPI_SETTOOLTIPANIMATION = 0x1017,
            SPI_GETTOOLTIPFADE = 0x1018,
            SPI_SETTOOLTIPFADE = 0x1019,
            SPI_GETCURSORSHADOW = 0x101A,
            SPI_SETCURSORSHADOW = 0x101B,

            //#if(_WIN32_WINNT >= =&H0501)
            SPI_GETMOUSESONAR = 0x101C,
            SPI_SETMOUSESONAR = 0x101D,
            SPI_GETMOUSECLICKLOCK = 0x101E,
            SPI_SETMOUSECLICKLOCK = 0x101F,
            SPI_GETMOUSEVANISH = 0x1020,
            SPI_SETMOUSEVANISH = 0x1021,
            SPI_GETFLATMENU = 0x1022,
            SPI_SETFLATMENU = 0x1023,
            SPI_GETDROPSHADOW = 0x1024,
            SPI_SETDROPSHADOW = 0x1025,
            SPI_GETBLOCKSENDINPUTRESETS = 0x1026,
            SPI_SETBLOCKSENDINPUTRESETS = 0x1027,

            //#endif /* _WIN32_WINNT >= =&H0501 */
            SPI_GETUIEFFECTS = 0x103E,
            SPI_SETUIEFFECTS = 0x103F,

            SPI_GETFOREGROUNDLOCKTIMEOUT = 0x2000,
            SPI_SETFOREGROUNDLOCKTIMEOUT = 0x2001,
            SPI_GETACTIVEWNDTRKTIMEOUT = 0x2002,
            SPI_SETACTIVEWNDTRKTIMEOUT = 0x2003,
            SPI_GETFOREGROUNDFLASHCOUNT = 0x2004,
            SPI_SETFOREGROUNDFLASHCOUNT = 0x2005,
            SPI_GETCARETWIDTH = 0x2006,
            SPI_SETCARETWIDTH = 0x2007,

            //#if(_WIN32_WINNT >= =&H0501)
            SPI_GETMOUSECLICKLOCKTIME = 0x2008,
            SPI_SETMOUSECLICKLOCKTIME = 0x2009,
            SPI_GETFONTSMOOTHINGTYPE = 0x200A,

            SPI_SETFONTSMOOTHINGTYPE = 0x200B
            //#endif /* WINVER >= =&H0500 */
        }


        [Flags]
        public enum TextAlignFlags
        {
            TA_LEFT = 0,
            TA_TOP = 0,
            TA_NOUPDATECP = 0,
            TA_UPDATECP = 1,
            TA_RIGHT = 2,
            TA_CENTER = 6,
            TA_BOTTOM = 8,
            TA_BASELINE = 24,
            TA_MASK = TA_BASELINE + TA_CENTER + TA_UPDATECP,
            TA_RTLREADING = 256,
            Vertical = 0x1000 //Custom
        }


        [Flags]
        public enum WindowExStyle
        {
            WS_EX_DLGMODALFRAME = 0x1,
            WS_EX_NOPARENTNOTIFY = 0x4,
            WS_EX_TOPMOST = 0x8,
            WS_EX_ACCEPTFILES = 0x10,
            WS_EX_TRANSPARENT = 0x20,
            WS_EX_MDICHILD = 0x40,
            WS_EX_TOOLWINDOW = 0x80,
            WS_EX_WINDOWEDGE = 0x100,
            WS_EX_CLIENTEDGE = 0x200,
            WS_EX_CONTEXTHELP = 0x400,
            WS_EX_RIGHT = 0x1000,
            WS_EX_LEFT = 0x0,
            WS_EX_RTLREADING = 0x2000,
            WS_EX_LTRREADING = 0x0,
            WS_EX_LEFTSCROLLBAR = 0x4000,
            WS_EX_RIGHTSCROLLBAR = 0x0,
            WS_EX_CONTROLPARENT = 0x10000,
            WS_EX_STATICEDGE = 0x20000,
            WS_EX_APPWINDOW = 0x40000,
            WS_EX_OVERLAPPEDWINDOW = WS_EX_WINDOWEDGE | WS_EX_CLIENTEDGE,
            WS_EX_PALETTEWINDOW = WS_EX_WINDOWEDGE | WS_EX_TOOLWINDOW | WS_EX_TOPMOST,
            WS_EX_LAYERED = 0x80000,
            WS_EX_NOINHERITLAYOUT = 0x100000,
            WS_EX_LAYOUTRTL = 0x400000,
            WS_EX_COMPOSITED = 0x2000000,
            WS_EX_NOACTIVATE = 0x8000000
        }


        [Flags]
        public enum WindowStyle
        {
            WS_NONE = 0x0,
            WS_OVERLAPPED = 0x0,
            WS_POPUP = -2147483648,
            WS_CHILD = 0x40000000,
            WS_MINIMIZE = 0x20000000,
            WS_VISIBLE = 0x10000000,
            WS_DISABLED = 0x8000000,
            WS_CLIPSIBLINGS = 0x4000000,
            WS_CLIPCHILDREN = 0x2000000,
            WS_MAXIMIZE = 0x1000000,
            WS_CAPTION = 0xC00000,
            WS_BORDER = 0x800000,
            WS_DLGFRAME = 0x400000,
            WS_VSCROLL = 0x200000,
            WS_HSCROLL = 0x100000,
            WS_SYSMENU = 0x80000,
            WS_THICKFRAME = 0x40000,
            WS_GROUP = 0x20000,
            WS_TABSTOP = 0x10000,
            WS_MINIMIZEBOX = 0x20000,
            WS_MAXIMIZEBOX = 0x10000,
            WS_TILED = WS_OVERLAPPED,
            WS_ICONIC = WS_MINIMIZE,
            WS_SIZEBOX = WS_THICKFRAME,
            WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW,

            WS_OVERLAPPEDWINDOW =
                WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX,
            WS_POPUPWINDOW = WS_POPUP | WS_BORDER | WS_SYSMENU,
            WS_CHILDWINDOW = WS_CHILD
        }

        //Guard against dll's being loaded more than once due to case sensitivity
        //i.e user32 User32 User32.Dll are all valid but will load the same dll multiple times.
        private const string User32 = "user32";
        private const string Gdi32 = "gdi32";
        private const string Kernel32 = "kernel32";
        private const string UXtheme = "uxtheme";
        private const string ComCtl32 = "comctl32";
        private const string MsImg32 = "msimg32";

        private const string OleAcc = "oleacc";

        //private const string Shell32 = "shell32";
        //private const string Shlwapi = "shlwapi";
        private const string UserEnv = "userenv";
        private const string Ole32 = "ole32";
        private const string Shcore = "shcore";

        public const string SystemDialogClassName = "#32770";
        public const int GWL_WNDPROC = -4;


        public const int S_OK = 0;

        public const int WVR_VALIDRECTS = 0x400;
        public const int PRF_NONCLIENT = 0x2;

        public const int DCX_INTERSECTRGN = 0x80;
        public const int DCX_WINDOW = 0x1;
        public const int DCX_CACHE = 0x2;

        public const int LF_FACESIZE = 32;

        //Public Const WM_SETREDRAW As Int32 = &HB
        public const int HWND_MESSAGE = -3;

        public const int SIZE_RESTORED = 0;
        public const int SIZE_MINIMIZED = 1;
        public const int SIZE_MAXIMIZED = 2;
        public const int SIZE_MAXSHOW = 3;
        public const int SIZE_MAXHIDE = 4;

        public const int WA_ACTIVE = 1;
        public const int WA_CLICKACTIVE = 2;
        public const int WA_INACTIVE = 0;

        public const int GWL_STYLE = -16;
        public const int GWL_EXSTYLE = -20;
        public const long GWL_HWNDPARENT = -8;

        public const int ULW_COLORKEY = 0x1;
        public const int ULW_ALPHA = 0x2;
        public const int ULW_OPAQUE = 0x4;

        public const int HWND_BOTTOM = 1;
        public const int HWND_DESKTOP = 0;
        public const int HWND_NOTOPMOST = -2;
        public const int HWND_TOP = 0;
        public const int HWND_TOPMOST = -1;

        public const int GCW_ATOM = -32;


        public const int CWP_ALL = 0x0;
        public const int CWP_SKIPDISABLED = 0x2;
        public const int CWP_SKIPINVISIBLE = 0x1;
        public const int CWP_SKIPTRANSPARENT = 0x4;


        // BlendOp:
        public const int AC_SRC_OVER = 0x0;

        // AlphaFormat:
        public const int AC_SRC_ALPHA = 0x1;


        public const int OBJID_NATIVEOM = -16;


        public const int PT_LOCAL = 0;
        public const int PT_TEMPORARY = 1;
        public const int PT_ROAMING = 2;
        public const int PT_MANDATORY = 4;


        public const int PAGE_EXECUTE_READWRITE = 64;


        private static int _currentProcessId;


        public static int CurrentProcessId
        {
            get
            {
                if (_currentProcessId == 0)
                {
                    _currentProcessId = GetCurrentProcessId();
                }

                return _currentProcessId;
            }
        }

        [DllImport(User32)]
        internal static extern bool EnumDisplayMonitors(IntPtr hdc, IntPtr lprcClip,
                                                        MonitorEnumProc lpfnEnum, IntPtr dwData);

        [DllImport(Shcore, EntryPoint = "GetScaleFactorForMonitor")]
        private static extern int GetScaleFactorForMonitorAPI(IntPtr hMon, out int pScale);

        internal static int GetScaleFactorForMonitor(IntPtr hMon)
        {
            if (!OSVersionHelper.IsWindows8Point1OrGreater())
                return 100; // This is the scaleFactor as a percentage, so return 100, aka Primary Monitor


            if (GetScaleFactorForMonitorAPI(hMon, out var pScale) != 0)
                return 100; // This is the scaleFactor as a percentage, so return 100, aka Primary Monitor

            return pScale;
        }

        [DllImport(User32, ExactSpelling = true)]
        public static extern IntPtr MonitorFromPoint(Point pt, int flags);

        [DllImport(User32, ExactSpelling = true)]
        public static extern IntPtr MonitorFromWindow(HandleRef handle, int flags);

        [DllImport(User32, ExactSpelling = true)]
        private static extern IntPtr MonitorFromRect([In] ref RECT lPrc, uint dwFlags);

        [DllImport(Shcore)]
        internal static extern uint GetDpiForMonitor(IntPtr hmonitor, uint dpiType, out int dpiX, out int dpiY);

        public static IntPtr MonitorFromRect(Rectangle rect, uint dwFlags)
        {
            var rc = new RECT(rect);
            return MonitorFromRect(ref rc, dwFlags);
        }

        public static int MakeLParam(Point pt)
        {
            return MakeLParam(Convert.ToInt16(pt.X), Convert.ToInt16(pt.Y));
        }

        public static int MakeLParam(short loWord, short hiWord)
        {
            return (Convert.ToInt32(hiWord) << 16) | (Convert.ToInt32(loWord) & short.MaxValue);
        }

        //Public Shared Function MakeLParamPtr(ByVal loWord As Int16, ByVal hiWord As Int16) As IntPtr
        //	Return New IntPtr(MakeLParam(loWord, hiWord))
        //End Function


        //Public Shared Function MakeWParam(ByVal loWord As Int16, ByVal hiWord As Int16) As IntPtr
        //	Return New IntPtr(MakeLParam(loWord, hiWord))
        //End Function


        public static short LoWord(int word)
        {
            return Convert.ToInt16(word & 0xffff);
        }

        public static short LoWord(IntPtr word)
        {
            return LoWord(word.ToInt32());
        }


        public static short HiWord(IntPtr word)
        {
            return HiWord(word.ToInt32());
        }

        public static short HiWord(int word)
        {
            return (short)((word >> 0x10) & 0xffff);
        }


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool GetGUIThreadInfo(int idThread, out GUITHREADINFO lpgui);


        [DllImport(ComCtl32)]
        public static extern int ImageList_DrawEx(IntPtr hIml, int i, IntPtr hdcDst, int x, int y, int dx, int dy,
                                                  int rgbBk, int rgbFg, DrawStyleFlags StyleFlags);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32, CallingConvention = CallingConvention.StdCall)]
        public static extern bool ShowWindow(IntPtr hWnd, ShowWindowFlags nCmdShow);


        [DllImport(Gdi32)]
        public static extern int SetBkColor(IntPtr hDc, int crColor);


        [DllImport(Gdi32)]
        public static extern IntPtr SelectObject(IntPtr hDc, IntPtr hObject);


        [DllImport(Gdi32)]
        public static extern int DeleteObject(IntPtr hObject);


        [DllImport(Gdi32)]
        public static extern int SetTextColor(IntPtr hDc, int crColor);


        [DllImport(Gdi32)]
        public static extern TextAlignFlags SetTextAlign(IntPtr hDc, TextAlignFlags wFlags);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(Gdi32, CharSet = CharSet.Auto, BestFitMapping = false, ThrowOnUnmappableChar = true)]
        public static extern bool GetTextExtentPoint32(IntPtr hDc, string lpsz, int cbString, ref Size lpSize);


        [DllImport(User32)]
        public static extern IntPtr GetWindowDC(IntPtr hWnd);


        [DllImport(User32)]
        public static extern int DrawFrameControl(IntPtr hDc, ref RECT lpRect, DrawFrameControlType uType,
                                                  DrawFrameControlState uState);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32, SetLastError = true)]
        public static extern bool SetProcessDPIAware();

        [DllImport(Kernel32)]
        public static extern int MulDiv(int nNumber, int nNumerator, int nDenominator);


        [DllImport(User32)]
        public static extern int SetFocus(IntPtr hWnd);

        [DllImport(User32)]
        public static extern IntPtr GetWindow(IntPtr hWnd, GetWindowCommand wCmd);


        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage")]
        [DllImport(User32)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage")]
        [DllImport(User32)]
        public static extern IntPtr GetForegroundWindow();


        [DllImport(User32)]
        public static extern IntPtr GetDCEx(IntPtr hWnd, IntPtr hrgnclip, int fdwOptions);


        [DllImport(User32)]
        public static extern IntPtr GetDC(IntPtr hWnd);


        [DllImport(Gdi32)]
        public static extern IntPtr CreateCompatibleDC(IntPtr hDC);

        [DllImport(Gdi32)]
        public static extern IntPtr CreateCompatibleBitmap(IntPtr hObject, int width, int height);

        [DllImport(MsImg32)]
        public static extern bool TransparentBlt(IntPtr hObject, int dstX, int dstY, int dstW, int nHeight,
                                                 IntPtr hObjSource, int srcX, int srcY, int srcW, int srcH, int color);

        [DllImport(Gdi32)]
        public static extern bool BitBlt(IntPtr hObject, int nXDest, int nYDest, int nWidth, int nHeight,
                                         IntPtr hObjSource, int nXSrc, int nYSrc, uint dwRop);

        [DllImport(User32, CharSet = CharSet.Auto)]
        public static extern int FillRect(IntPtr hDC, ref RECT rect, IntPtr hBrush);

        [DllImport(Gdi32, CharSet = CharSet.Auto)]
        public static extern IntPtr CreateSolidBrush(uint crColor);

        [DllImport("olepro32.dll", CharSet = CharSet.Auto)]
        public static extern void OleCreatePictureIndirect([MarshalAs(UnmanagedType.Struct)] ref PICTDESC pPictDesc,
                                                           ref Guid riid, bool fOwn, ref object ppvObj);

        [DllImport(Gdi32)]
        public static extern bool DeleteDC(IntPtr hdc);


        [DllImport(User32)]
        public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDc);


        [DllImport(User32, CharSet = CharSet.Auto)]
        public static extern int RegisterWindowMessage(string msg);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy,
                                               SetWindowPosFlags wFlags);


        [DllImport(User32)]
        public static extern IntPtr BeginDeferWindowPos(int nNumWindows);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool EndDeferWindowPos(IntPtr hWinPosInfo);


        [DllImport(User32)]
        public static extern IntPtr DeferWindowPos(IntPtr hWinPosInfo, IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y,
                                                   int cx, int cy, SetWindowPosFlags flags);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32, CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool AdjustWindowRectEx(ref RECT lpRect, int dwStyle,
                                                     [MarshalAs(UnmanagedType.Bool)] bool bMenu, int dwExStyle);

        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage")]
        public static Rectangle AdjustWindowRectEx(Rectangle bounds, int dwStyle, bool bMenu, int dwExStyle)
        {
            var rc = new RECT(bounds);
            if (AdjustWindowRectEx(ref rc, dwStyle, bMenu, dwExStyle))
            {
                return rc.ToRectangle();
            }


            return new Rectangle();
        }

        public static Rectangle AdjustWindowRectEx(HandleRef hWnd, Rectangle bounds, bool bMenu)
        {
            var rc = new RECT(bounds);
            var style = GetWindowLong(hWnd, GWL_STYLE).ToInt32();
            var styleEx = GetWindowLong(hWnd, GWL_EXSTYLE).ToInt32();
            if (AdjustWindowRectEx(ref rc, style, bMenu, styleEx))
            {
                return rc.ToRectangle();
            }


            return new Rectangle();
        }


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool RedrawWindow(IntPtr hWnd, [In] ref Rectangle lprcUpdate, IntPtr hrgnUpdate,
                                               RedrawWindowFlags fuRedraw);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool RedrawWindow(IntPtr hWnd, ref IntPtr lprcUpdate, IntPtr hrgnUpdate,
                                               RedrawWindowFlags fuRedraw);


        [DllImport(User32)]
        public static extern IntPtr WindowFromPoint(int x, int y);


        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "1")]
        [DllImport(User32)]
        public static extern IntPtr ChildWindowFromPoint(IntPtr hWnd, [In] ref Point pt);

        [DllImport(User32)]
        public static extern IntPtr ChildWindowFromPointEx(IntPtr hWndParent, int ptx, int pty, int uFlags);


        [DllImport(User32)]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);


        [DllImport(Kernel32)]
        private static extern int GetCurrentProcessId();


        [SuppressMessage("Microsoft.Usage", "CA1806:DoNotIgnoreMethodResults",
            MessageId = "LevitJames.Libraries.NativeMethods.GetWindowThreadProcessId(System.IntPtr,System.Int32@)")]
        public static bool IsWindowInCurrentProcess(IntPtr hWnd)
        {
            int processID;
            GetWindowThreadProcessId(hWnd, out processID);
            return processID == CurrentProcessId;
        }


        [DllImport(User32)]
        public static extern int ValidateRect(IntPtr hWnd, IntPtr lpRect);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool InvalidateRect(IntPtr hWnd, IntPtr lpRect, int bErase);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool InvalidateRect(IntPtr hWnd, [In] ref RECT lpRect, int bErase);

        public static bool InvalidateRect(IntPtr hWnd, Rectangle lpRect, int bErase)
        {
            var rc = new RECT(lpRect);
            return InvalidateRect(hWnd, ref rc, bErase);
        }


        [DllImport(Gdi32)]
        public static extern int SetWindowOrgEx(IntPtr hDc, int nX, int nY, out Point lpPoint);


        [DllImport(User32)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int uMsg, IntPtr wParam, IntPtr lParam);

        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return")]
        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "3")]
        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "2")]
        [DllImport(User32)]
        public static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, int lParam);

        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return")]
        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "2")]
        [DllImport(User32)]
        public static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, [In] ref Point lParam);

        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return")]
        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "2")]
        [DllImport(User32)]
        public static extern IntPtr SendMessage(HandleRef hWnd, int uMsg, IntPtr wParam, IntPtr lParam);


        [DllImport(User32, CallingConvention = CallingConvention.StdCall)]
        public static extern IntPtr GetActiveWindow();


        [DllImport(User32)]
        public static extern int PostMessage(IntPtr hWnd, int uMsg, IntPtr wParam, IntPtr lParam);

        [DllImport(User32)]
        public static extern int PostMessage(IntPtr hWnd, int uMsg, int wParam, int lParam);


        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return")]
        [DllImport(User32, CharSet = CharSet.Auto)]
        public static extern IntPtr CallNextHookEx(WindowsHookSafeHandle hook, int nCode, IntPtr wParam, IntPtr lParam);


        [DllImport(Kernel32, CharSet = CharSet.Auto)]
        public static extern IntPtr GetCommandLine();


        [DllImport(User32)]
        public static extern IntPtr GetFocus();


        [DllImport(User32, CharSet = CharSet.Auto, BestFitMapping = false, ThrowOnUnmappableChar = true)]
        public static extern int GetClassName(IntPtr hWnd, [MarshalAs(UnmanagedType.LPTStr)] StringBuilder lpClassName,
                                              int nMaxCount);

        public static string GetClassName(IntPtr hWnd)
        {
            var className = new StringBuilder(capacity: 255);
            var chars = GetClassName(hWnd, className, nMaxCount: 255);
            if (chars > 0)
            {
                return className.ToString(startIndex: 0, length: chars);
            }

            return null;
        }


        [DllImport(User32)]
        public static extern int GetClassLong(IntPtr hWnd, int nIndex);


        [DllImport(User32)]
        public static extern IntPtr GetDlgItem(HandleRef hWnd, int item);


        [DebuggerStepThrough]
        public static Atom GetClassAtom(IntPtr hWnd)
        {
            return new Atom(GetClassLong(hWnd, GCW_ATOM));
        }


        public static Atom GetClassAtomFromName(string windowClassName)
        {
            return RegisterClipboardFormat(windowClassName);
        }


        [DllImport(User32, CharSet = CharSet.Auto, BestFitMapping = false, ThrowOnUnmappableChar = true)]
        public static extern Atom RegisterClipboardFormat(string lpszFormat);


        [DllImport(User32, CharSet = CharSet.Auto, BestFitMapping = false, ThrowOnUnmappableChar = true)]
        public static extern int GetWindowText(IntPtr hWnd, [MarshalAs(UnmanagedType.LPTStr)] StringBuilder lpString,
                                               int cch);

        public static string GetWindowText(IntPtr hWnd)
        {
            var text = new StringBuilder(capacity: 255);
            var chars = GetWindowText(hWnd, text, cch: 255);
            if (chars > 0)
            {
                return text.ToString(startIndex: 0, length: chars);
            }

            return null;
        }


        [DllImport(User32, CharSet = CharSet.Auto, BestFitMapping = false, ThrowOnUnmappableChar = true)]
        public static extern IntPtr FindWindowEx(IntPtr hWndParent, IntPtr hWndChildAfter, string lpszClass,
                                                 string lpszWindow);

        [DllImport(User32, CharSet = CharSet.Auto, BestFitMapping = false, ThrowOnUnmappableChar = true)]
        public static extern IntPtr FindWindowEx(IntPtr hWndParent, IntPtr hWndChildAfter, Atom lpszClass,
                                                 string lpszWindow);


        [DllImport(User32)]
        public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);


        [DllImport(User32)]
        public static extern IntPtr GetParent(IntPtr hWnd);


        [DllImport(User32)]
        public static extern int UpdateWindow(IntPtr hWnd);


        [DllImport(User32, ExactSpelling = true)]
        public static extern int UpdateLayeredWindow(IntPtr hwnd, IntPtr hdcDst, ref Point pptDst, ref Size psize,
                                                     IntPtr hdcSrc, ref Point pprSrc, int crKey,
                                                     ref BLENDFUNCTION pblend, int dwFlags);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool GetWindowRect(IntPtr hWnd, ref RECT lpRect);

        public static Rectangle GetWindowRect(IntPtr hWnd)
        {
            var rc = new RECT();
            if (GetWindowRect(hWnd, ref rc))
            {
                return new Rectangle(rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top);
            }

            return new Rectangle();
        }


        //Note this only does one rectangle at the momement
        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        private static extern bool MapWindowPoints(IntPtr hWndFrom, IntPtr hWndTo, ref RECT lpRect, int cPoints);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        private static extern bool MapWindowPoints(IntPtr hWndFrom, IntPtr hWndTo, ref Point lpPoint, int cPoints);

        public static Rectangle MapWindowPoints(IntPtr hWndFrom, IntPtr hWndTo, Rectangle rc)
        {
            var rcWin32 = new RECT(rc);
            if (MapWindowPoints(hWndFrom, hWndTo, ref rcWin32, cPoints: 2)) //2 as there are two points in a rect
            {
                return new Rectangle(rcWin32.Left, rcWin32.Top, rcWin32.Right - rcWin32.Left,
                                     rcWin32.Bottom - rcWin32.Top);
            }

            return new Rectangle();
        }

        public static Point MapWindowPoints(IntPtr hWndFrom, IntPtr hWndTo, Point pt)
        {
            if (MapWindowPoints(hWndFrom, hWndTo, ref pt, cPoints: 1))
            {
                return pt;
            }

            return new Point();
        }


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool GetClientRect(IntPtr hWnd, ref RECT lpRect);

        public static Rectangle GetClientRect(IntPtr hWnd)
        {
            var rc = new RECT();
            if (GetClientRect(hWnd, ref rc))
            {
                return new Rectangle(rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top);
            }

            return new Rectangle();
        }


        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return")]
        [DllImport(User32, EntryPoint = "GetWindowLong", CharSet = CharSet.Auto)]
        private static extern IntPtr GetWindowLongx86(HandleRef hWnd, int nIndex);

        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "2")]
        [SuppressMessage("Microsoft.Interoperability", "CA1400:PInvokeEntryPointsShouldExist")]
        [DllImport(User32, EntryPoint = "GetWindowLongPtr", CharSet = CharSet.Auto)]
        private static extern IntPtr GetWindowLongX64(HandleRef hWnd, int nIndex);

        public static IntPtr GetWindowLong(HandleRef hWnd, int nIndex)
        {
            if (IntPtr.Size == 4)
            {
                return GetWindowLongx86(hWnd, nIndex);
            }

            return GetWindowLongX64(hWnd, nIndex);
        }


        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return")]
        [DllImport(User32, EntryPoint = "SetWindowLong")]
        private static extern IntPtr SetWindowLongx86(HandleRef hWnd, int nIndex, int dwNewLong);

        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return")]
        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "2")]
        [SuppressMessage("Microsoft.Interoperability", "CA1400:PInvokeEntryPointsShouldExist")]
        [DllImport(User32, EntryPoint = "SetWindowLongPtr")]
        private static extern IntPtr SetWindowLongX64(HandleRef hWnd, int nIndex, int dwNewLong);

        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return")]
        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "2")]
        [SuppressMessage("Microsoft.Interoperability", "CA1400:PInvokeEntryPointsShouldExist")]
        [DllImport(User32, EntryPoint = "SetWindowLong")]
        private static extern IntPtr SetWindowLongx86(HandleRef hWnd, int nIndex, HandleRef dwNewLong);


        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "2")]
        [SuppressMessage("Microsoft.Interoperability", "CA1400:PInvokeEntryPointsShouldExist")]
        [DllImport(User32, EntryPoint = "SetWindowLongPtr")]
        private static extern IntPtr SetWindowLongX64(HandleRef hWnd, int nIndex, HandleRef dwNewLong);


        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return")]
        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "2")]
        [SuppressMessage("Microsoft.Interoperability", "CA1400:PInvokeEntryPointsShouldExist")]
        [DllImport(User32, EntryPoint = "SetWindowLong")]
        private static extern IntPtr SetWindowLongx86(HandleRef hWnd, int nIndex, IntPtr dwNewLong);

        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "2")]
        [SuppressMessage("Microsoft.Interoperability", "CA1400:PInvokeEntryPointsShouldExist")]
        [DllImport(User32, EntryPoint = "SetWindowLongPtr")]
        private static extern IntPtr SetWindowLongX64(HandleRef hWnd, int nIndex, IntPtr dwNewLong);

        public static IntPtr SetWindowLong(HandleRef hWnd, int nIndex, int dwNewLong)
        {
            if (IntPtr.Size == 4)
            {
                return SetWindowLongx86(hWnd, nIndex, dwNewLong);
            }

            return SetWindowLongX64(hWnd, nIndex, dwNewLong);
        }

        public static IntPtr SetWindowLong(HandleRef hWnd, int nIndex, HandleRef dwNewLong)
        {
            if (IntPtr.Size == 4)
            {
                return SetWindowLongx86(hWnd, nIndex, dwNewLong);
            }

            return SetWindowLongX64(hWnd, nIndex, dwNewLong);
        }

        public static IntPtr SetWindowLong(HandleRef hWnd, int nIndex, IntPtr dwNewLong)
        {
            if (IntPtr.Size == 4)
            {
                return SetWindowLongx86(hWnd, nIndex, dwNewLong);
            }

            return SetWindowLongX64(hWnd, nIndex, dwNewLong);
        }

        [DllImport(User32, EntryPoint = "GetWindowInfo")]
        private static extern bool GetWindowInfo(IntPtr hWnd, out WINDOWINFOSTRUCT info);

        public static WINDOWINFOSTRUCT GetWindowInfo(IntPtr hWnd)
        {
            GetWindowInfo(hWnd, out var wInfo);
            return wInfo;
        }


        [DllImport(User32, CharSet = CharSet.Auto)]
        public static extern IntPtr CallWindowProc(IntPtr lpPrevWndFunc, IntPtr hwnd, Int32 msg, IntPtr wParam, IntPtr lParam);


        [DllImport(ComCtl32)]
        internal static extern Int32 SetWindowSubclass(IntPtr hWnd, SubClassProcDelegate newProc, UIntPtr uIdSubclass, IntPtr dwRefData);


        [DllImport(ComCtl32)]
        internal static extern Int32 RemoveWindowSubclass(IntPtr hWnd, SubClassProcDelegate newProc, UIntPtr uIdSubclass);


        [DllImport(ComCtl32)]
        internal static extern IntPtr DefSubclassProc(IntPtr hWnd, Int32 Msg, IntPtr wParam, IntPtr lParam);


        [DllImport(Kernel32, CharSet = CharSet.Auto)]
        public static extern int GetModuleFileName(IntPtr hModule, StringBuilder lpFileName, int nSize);


        [DllImport(Kernel32, CharSet = CharSet.Auto)]
        public static extern IntPtr GetModuleHandle(string lpModuleName);


        [DllImport(Kernel32)]
        public static extern int GetCurrentThreadId();


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool IsWindow(IntPtr hWnd);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool ShowOwnedPopups(IntPtr hWnd, [MarshalAs(UnmanagedType.Bool)] bool fShow);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool IsIconic(IntPtr hWnd);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool IsZoomed(IntPtr hWnd);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool IsWindowVisible(IntPtr hWnd);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool IsWindowEnabled(IntPtr hWnd);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool EnableWindow(IntPtr hWnd, [MarshalAs(UnmanagedType.Bool)] bool fEnable);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool EnumChildWindows(IntPtr hWndParent, EnumChildProc lpEnumFunc, IntPtr lParam);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool EnumThreadWindows(int dwThreadID, EnumThreadWndProc lpfn, IntPtr lParam);

        [DllImport(User32)]
        public static extern IntPtr GetAncestor(IntPtr hWnd, GAFlags gaFlags);


        [DllImport(User32)]
        public static extern int SetActiveWindow(IntPtr hWnd);

        [DllImport(MsImg32)]
        public static extern int AlphaBlend(IntPtr hdcDest, int nXOriginDest, int nYOriginDest, int nWidthDest,
                                            int nHeightDest, IntPtr hdcSrc, int nXOriginSrc, int nYOriginSrc,
                                            int nWidthSrc, int nHeightSrc, BLENDFUNCTION lBlendFunction);


        public static void ThrowWin32Error()
        {
            ThrowWin32Error(message: null);
        }

        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        [SuppressMessage("Microsoft.Usage", "CA2201:DoNotRaiseReservedExceptionTypes")]
        public static void ThrowWin32Error(string message)
        {
            var er = Marshal.GetLastWin32Error();
            if (er != 0)
            {
                if (string.IsNullOrEmpty(message))
                {
                    throw new Win32Exception(er);
                }

                throw new Win32Exception(er, message);
            }
        }


        [DllImport(OleAcc)]
        public static extern int AccessibleObjectFromPoint(Point ptScreen, out IAccessible ppvObject,
                                                           out object pvarChild);

        [DllImport(OleAcc)]
        public static extern int AccessibleObjectFromWindow(IntPtr hWnd, int dwid, [In] ref Guid riid,
                                                            [MarshalAs(UnmanagedType.IDispatch)] ref object ppvobject);


        [DllImport(OleAcc)]
        private static extern int AccessibleChildren(IAccessible element, int iChildStart, int cChildren, [Out] object[] rgvarChildren, out int pcObtained);

        public static object[] AccessibleChildren(IAccessible element)
        {
            var children = new object[element.accChildCount];
            if (element.accChildCount > 0)
            {
                AccessibleChildren(element, iChildStart: 0, cChildren: children.Length, rgvarChildren: children,
                                   pcObtained: out _);
            }

            return children;
        }


        //public static string RoleText(AccessibleRole role)
        //{
        //    StringBuilder sb;
        //    var roleTextLength = GetRoleText((int) role, out sb, cchRoleMax: 0);
        //    GetRoleText((int) role, out sb, roleTextLength + 1);
        //    return sb.ToString();
        //}

        [DllImport(OleAcc)]
        private static extern int GetRoleText(int dwRole, out StringBuilder lpszRole, int cchRoleMax);


        [DllImport(OleAcc)]
        public static extern int WindowFromAccessibleObject(IAccessible pAcc, ref IntPtr hwnd);


        [DllImport(UXtheme, CharSet = CharSet.Unicode)]
        public static extern int GetCurrentThemeName(StringBuilder pszThemeFileName, int dwMaxNameChars,
                                                     StringBuilder pszColorBuff, int dwMaxColorChars,
                                                     StringBuilder pszSizeBuff, int cchMaxSizeChars);


        [SuppressMessage("Microsoft.Usage", "CA1806:DoNotIgnoreMethodResults",
            MessageId =
                "LevitJames.Libraries.NativeMethods.GetCurrentThemeName(System.Text.StringBuilder,System.Int32,System.Text.StringBuilder,System.Int32,System.Text.StringBuilder,System.Int32)"
        )]
        //public static string GetCurrentThemeFilename()
        //{
        //    if (VisualStyleInformation.IsEnabledByUser)
        //    {
        //        var pszThemeFileName = new StringBuilder(capacity: 512);
        //        GetCurrentThemeName(pszThemeFileName, pszThemeFileName.Capacity, pszColorBuff: null, dwMaxColorChars: 0,
        //                            pszSizeBuff: null, cchMaxSizeChars: 0);
        //        return pszThemeFileName.ToString();
        //    }

        //    return string.Empty;
        //}

        [DllImport(UserEnv)]
        public static extern bool GetProfileType(ref int pdwflags);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(Kernel32, SetLastError = true, CallingConvention = CallingConvention.Winapi)]
        private static extern bool IsWow64Process(IntPtr hProcess, out bool wow64Process);

        public static bool IsWow64Process()
        {
            bool wow64Process;
            IsWow64Process(Process.GetCurrentProcess().Handle, out wow64Process);
            return wow64Process;
        }


        public static bool Is64BitOS()
        {
            if (IntPtr.Size == 8 || (IntPtr.Size == 4 && IsWow64Process()))
            {
                return true;
            }

            return false;
        }


        [DllImport(Ole32)]
        public static extern int CoRegisterMessageFilter(IMessageFilter lpMessageFilter,
                                                         out IMessageFilter lplpMessageFilter);

        [DllImport(Kernel32)]
        public static extern bool VirtualProtect(IntPtr lpAddress, int dwSize, int flNewProtect, ref int lpflOldProtect);


        [DllImport(Kernel32)]
        public static extern void CopyMemory(IntPtr dest, IntPtr src, int count);

        [DllImport(Kernel32, EntryPoint = "CopyMemory")]
        public static extern void CopyMemoryDestByRef(ref IntPtr dest, IntPtr src, int count);

        [DllImport(Kernel32, EntryPoint = "CopyMemory")]
        public static extern void CopyMemorySrcByRef(IntPtr dest, ref IntPtr src, int count);


        internal delegate bool MonitorEnumProc(IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData);


        [StructLayout(LayoutKind.Sequential)]
        public struct NCCALCSIZE_PARAMS
        {
            public RECT rgrc0;
            public RECT rgrc1;
            public RECT rgrc2;
            public IntPtr lppos;
        }


        [StructLayout(LayoutKind.Sequential)]
        public class MINMAXINFO
        {
            public Point ptReserved;
            public Point ptMaxSize;
            public Point ptMaxPosition;
            public Point ptMinTrackSize;
            public Point ptMaxTrackSize;
        }


        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct LOGFONT
        {
            public int lfHeight;
            public int lfWidth;
            public int lfEscapement;
            public int lfOrientation;
            public int lfWeight;
            public byte lfItalic;
            public byte lfUnderline;
            public byte lfStrikeOut;
            public byte lfCharSet;
            public byte lfOutPrecision;
            public byte lfClipPrecision;
            public byte lfQuality;
            public byte lfPitchAndFamily;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = LF_FACESIZE)]
            public string lfFaceName;
        }


        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;

            public RECT(int left, int top, int width, int height) : this()
            {
                Left = left;
                Top = top;
                Right = left + width;
                Bottom = top + height;
            }

            public RECT(Rectangle rc) : this()
            {
                Left = rc.Left;
                Top = rc.Top;
                Right = rc.Right;
                Bottom = rc.Bottom;
            }

            public Rectangle ToRectangle()
            {
                return Rectangle.FromLTRB(Left, Top, Right, Bottom);
            }

            public int Width
            {
                get => Right - Left;
                set => Right = Left + value;
            }

            public int Height
            {
                get => Bottom - Top;
                set => Bottom = Top + value;
            }

            public void Offset(int x, int y)
            {
                if (x != 0)
                {
                    Left += x;
                    Right += x;
                }

                if (y != 0)
                {
                    Top += y;
                    Bottom += y;
                }
            }

            public void Inflate(int x, int y)
            {
                if (x != 0)
                {
                    Left -= x;
                    Right += x;
                }

                if (y != 0)
                {
                    Top -= y;
                    Bottom += y;
                }
            }
        }


        [StructLayout(LayoutKind.Sequential)]
        public struct STYLESTRUCT
        {
            public int styleOld;
            public int styleNew;
        }


        [StructLayout(LayoutKind.Sequential)]
        public struct WINDOWPOS
        {
            public IntPtr hwnd;
            public IntPtr hwndInsertAfter;
            public int x;
            public int y;
            public int cx;
            public int cy;
            public SetWindowPosFlags flags;
        }


        [StructLayout(LayoutKind.Sequential)]
        public struct EVENTMSG
        {
            public int message;
            public IntPtr paramL;
            public IntPtr paramH;
            public int time;
            public IntPtr hWnd;
        }


        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct BLENDFUNCTION
        {
            public byte BlendOp;
            public byte BlendFlags;
            public byte SourceConstantAlpha;
            public byte AlphaFormat;

            public BLENDFUNCTION(byte BlendOp, byte BlendFlags, byte SourceConstantAlpha, byte AlphaFormat) : this()
            {
                this.BlendOp = BlendOp;
                this.BlendFlags = BlendFlags;
                this.SourceConstantAlpha = SourceConstantAlpha;
                this.AlphaFormat = AlphaFormat;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct GUITHREADINFO
        {
            public int cbSize;
            public GUITHREADINFOFlags flags;
            public IntPtr hwndActive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;
            public IntPtr hwndMoveSize;
            public IntPtr hwndCaret;
            public RECT rcCaret;

            public static GUITHREADINFO NewGUITHREADINFO()
            {
                var guiTHREADINFO = new GUITHREADINFO();
                guiTHREADINFO.cbSize = Marshal.SizeOf(typeof(GUITHREADINFO));
                return guiTHREADINFO;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct PICTDESC
        {
            public int cbSizeOfStruct;
            public int picType;
            public int handle1;
            public int handle2;
            public int handle3;
        }


        [DebuggerStepThrough]
        [StructLayout(LayoutKind.Explicit, Size = 4)]
        public struct Atom : IComparable, IFormattable, IConvertible, IComparable<int>, IEquatable<int>
        {
            public const int MaxValue = 2147483647;
            public const int MinValue = -2147483648;

            [FieldOffset(offset: 0)] public readonly int _value;

            public static Atom Zero;

            public Atom(int value) : this()
            {
                _value = value;
            }

            public int CompareTo(object value)
            {
                if (value == null)
                {
                    return 1;
                }

                var num = Convert.ToInt32(value);
                if (_value < num)
                {
                    return -1;
                }

                if (_value > num)
                {
                    return 1;
                }

                return 0;
            }

            public int CompareTo(int value)
            {
                if (_value < value)
                {
                    return -1;
                }

                if (_value > value)
                {
                    return 1;
                }

                return 0;
            }

            public override bool Equals(object obj)
            {
                return obj is int && (_value == Convert.ToInt32(obj));
            }

            public bool Equals(int obj)
            {
                return _value == obj;
            }

            public override int GetHashCode()
            {
                return _value;
            }

            public override string ToString()
            {
                return _value.ToString();
            }

            public string ToString(string format)
            {
                return _value.ToString(format);
            }

            public string ToString(IFormatProvider provider)
            {
                return _value.ToString(provider);
            }

            public string ToString(string format, IFormatProvider provider)
            {
                return _value.ToString(format, provider);
            }

            public static Atom Parse(string s)
            {
                return new Atom(int.Parse(s));
            }

            public static Atom Parse(string s, NumberStyles style)
            {
                return new Atom(int.Parse(s, style));
            }

            public static Atom Parse(string s, IFormatProvider provider)
            {
                return new Atom(int.Parse(s, provider));
            }

            public static Atom Parse(string s, NumberStyles style, IFormatProvider provider)
            {
                return new Atom(int.Parse(s, style, provider));
            }

            public static bool TryParse(string s, out Atom result)
            {
                int value;
                var r = int.TryParse(s, out value);
                result = r ? new Atom(value) : default(Atom);

                return r;
            }

            public static bool TryParse(string s, NumberStyles style, IFormatProvider provider, out Atom result)
            {
                int value;
                var r = int.TryParse(s, style, provider, out value);
                result = r ? new Atom(value) : default(Atom);
                return r;
            }

            public TypeCode GetTypeCode()
            {
                return TypeCode.Int32;
            }

            bool IConvertible.ToBoolean(IFormatProvider provider)
            {
                return Convert.ToBoolean(_value);
            }


            char IConvertible.ToChar(IFormatProvider provider)
            {
                return Convert.ToChar(_value);
            }

            sbyte IConvertible.ToSByte(IFormatProvider provider)
            {
                return Convert.ToSByte(_value);
            }

            byte IConvertible.ToByte(IFormatProvider provider)
            {
                return Convert.ToByte(_value);
            }

            short IConvertible.ToInt16(IFormatProvider provider)
            {
                return Convert.ToInt16(_value);
            }

            ushort IConvertible.ToUInt16(IFormatProvider provider)
            {
                return Convert.ToUInt16(_value);
            }

            int IConvertible.ToInt32(IFormatProvider provider)
            {
                return _value;
            }

            uint IConvertible.ToUInt32(IFormatProvider provider)
            {
                return Convert.ToUInt32(_value);
            }

            long IConvertible.ToInt64(IFormatProvider provider)
            {
                return Convert.ToInt64(_value);
            }

            ulong IConvertible.ToUInt64(IFormatProvider provider)
            {
                return Convert.ToUInt64(_value);
            }

            float IConvertible.ToSingle(IFormatProvider provider)
            {
                return Convert.ToSingle(_value);
            }

            double IConvertible.ToDouble(IFormatProvider provider)
            {
                return Convert.ToDouble(_value);
            }

            decimal IConvertible.ToDecimal(IFormatProvider provider)
            {
                return Convert.ToDecimal(_value);
            }

            DateTime IConvertible.ToDateTime(IFormatProvider provider)
            {
                return ((IConvertible)_value).ToDateTime(provider);
            }

            object IConvertible.ToType(Type type, IFormatProvider provider)
            {
                return ((IConvertible)_value).ToType(type, provider);
            }

            public static explicit operator Atom(int value)
            {
                return new Atom(value);
            }

            public static explicit operator int(Atom value)
            {
                return Convert.ToInt32(value._value);
            }

            public static bool operator ==(Atom value1, Atom value2)
            {
                return value1._value == value2._value;
            }

            public static bool operator ==(Atom value1, int value2)
            {
                return value1._value == value2;
            }

            public static bool operator ==(int value1, Atom value2)
            {
                return value1 == value2._value;
            }

            public static bool operator !=(Atom value1, Atom value2)
            {
                return value1._value != value2._value;
            }

            public static bool operator !=(Atom value1, int value2)
            {
                return value1._value != value2;
            }

            public static bool operator !=(int value1, Atom value2)
            {
                return value1 != value2._value;
            }
        }


        [StructLayout(LayoutKind.Sequential)]
        public struct WINDOWINFOSTRUCT
        {
            public int cbSize;
            public RECT rcWindow;
            public RECT rcClient;
            public int dwStyle;
            public int dwExStyle;
            public int dwWindowStatus;
            public int cxWindowBorders;
            public int cyWindowBorders;
            public int atomWindowtype;
            public int wCreatorVersion;
        }

        internal delegate IntPtr SubClassProcDelegate(IntPtr hWnd, Int32 Msg, IntPtr wParam, IntPtr lParam, UIntPtr uIdSubclass, IntPtr dwRefData);


        public class WindowsHookSafeHandle : SafeHandle
        {
            // ReSharper disable once NotAccessedField.Local
            private fnHookProc _fnHookProc;
            //Must keep a ref to this so we don't get a CallbackOnCollectedDelegate Exception


            public WindowsHookSafeHandle() : base(IntPtr.Zero, ownsHandle: true) { }


            public override bool IsInvalid
            {
                [DebuggerStepThrough]
                get { return handle == IntPtr.Zero; }
            }

            [DllImport(User32, CharSet = CharSet.Auto)]
            private static extern int UnhookWindowsHookEx(IntPtr hook);


            [DllImport(User32, EntryPoint = "SetWindowsHookEx", SetLastError = true)]
            private static extern WindowsHookSafeHandle SetWindowsHookExInternal(int idHook, fnHookProc lpfn,
                                                                                 IntPtr hMod, int dwThreadId);


            protected override bool ReleaseHandle()
            {
                var retVal = UnhookWindowsHookEx(handle);
                if (retVal == 0)
                {
                    //Do not throw may be called from a finalizer which may result is a crash
                    //NativeMethods.ThrowWin32Error("StopHook Failed")
                }

                handle = IntPtr.Zero;
                _fnHookProc = null;
                return true;
            }


            public static WindowsHookSafeHandle SetWindowsHookEx(int idHook, fnHookProc lpfn, IntPtr hMod,
                                                                 int dwThreadId)
            {
                var safeHandle = SetWindowsHookExInternal(idHook, lpfn, hMod, dwThreadId);
                //Debug.WriteLine(idHook.ToString)
                if (!safeHandle.IsInvalid)
                {
                    safeHandle._fnHookProc = lpfn;
                }

                return safeHandle;
            }
        }

        internal static int IntPtrToInt32(IntPtr value)
        {
            if (4 == IntPtr.Size)
            {
                return value.ToInt32();
            }

            var longValue = (long)value;
            longValue = Math.Min(int.MaxValue, longValue);
            longValue = Math.Max(int.MinValue, longValue);
            return (int)longValue;
        }
    }
}