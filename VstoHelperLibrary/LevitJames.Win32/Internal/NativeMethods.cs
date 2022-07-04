// © Copyright 2018 Levit & James, Inc.

using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security;
using System.Text;

 #pragma warning disable 649

// ReSharper disable All

namespace LevitJames.Win32
{

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 4)]
    public class MONITORINFOEX
    {
        internal int cbSize = Marshal.SizeOf(typeof(MONITORINFOEX));
        public RectangleI rcMonitor = new RectangleI();
        public RectangleI rcWork = new RectangleI();
        public int dwFlags = 0;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 32)]
        public char[] szDevice = new char[32];
    }


    [Flags]
	public enum SetWindowPosFlags
	{
		None = 0x0,
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
		SWP_NOSENDCHANGING = 0x400,
		SWP_NOSIZE_NOMOVE = SWP_NOSIZE | SWP_NOMOVE,
		SWP_NOSIZE_NOMOVE_NOACTIVATE = SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE
	}

    [Flags]
    public enum WindowStyles
    {
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

    [Flags]
    public enum WindowExStyles
    {
	    None = 0x0,
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

    [GeneratedCode("", ""), SuppressUnmanagedCodeSecurity,
     SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
    [CompilerGenerated]
    internal static partial class NativeMethods
    {
	    //Guard against dll's being loaded more than once due to case sensitivity
	    //i.e user32 User32 User32.Dll are all valid but will load the same dll multiple times.
	    internal const string User32 = "user32";
	    private const string Gdi32 = "gdi32";
	    private const string Kernel32 = "kernel32";
	    private const string UxTheme = "uxtheme";
	    private const string ComCtl32 = "comctl32";
	    private const string MsImg32 = "msimg32";
	    private const string OleAcc = "oleacc";
	    private const string Shell32 = "shell32";
	    private const string Shlwapi = "shlwapi";
	    private const string GdiPlus = "gdiplus";
	    private const string Ole32 = "ole32";
	    private const string Uxtheme = "uxtheme";
	    private const string Shcore = "shcore";
	    private const string Comdlg32 = "comdlg32";
        private const string OleAut32 = "oleaut32";
        

        public delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);


        public delegate bool EnumThreadWndProc(IntPtr hWnd, IntPtr lParam);
 
        public delegate IntPtr fnHookProc(int nCode, IntPtr wParam, IntPtr lParam);

        public delegate void TimerProc(IntPtr hWnd, uint uMsg, IntPtr nIDEvent, uint dwTime);

        public const int ILD_TRANSPARENT = 0x00000001;
        public const int GWL_WNDPROC = -4;

        [ComImportAttribute()]
        [GuidAttribute("46EB5926-582E-4017-9FDF-E8998DAA0950")]
        [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IImageList {
            [PreserveSig]
            int Add(
            IntPtr hbmImage,
            IntPtr hbmMask,
            ref int pi);

            [PreserveSig]
            int ReplaceIcon(
            int i,
            IntPtr hicon,
            ref int pi);

            [PreserveSig]
            int SetOverlayImage(
            int iImage,
            int iOverlay);

            [PreserveSig]
            int Replace(
            int i,
            IntPtr hbmImage,
            IntPtr hbmMask);

            [PreserveSig]
            int AddMasked(
            IntPtr hbmImage,
            int crMask,
            ref int pi);

            [PreserveSig]
            int Draw(
            IntPtr pimldp);

            [PreserveSig]
            int Remove(
            int i);

            [PreserveSig]
            int GetIcon(
            int i,
            DrawStyleFlags flags,
            ref IntPtr picon);

            [PreserveSig]
            int GetImageInfo(
            int i,
            IntPtr pImageInfo);

            [PreserveSig]
            int Copy(
            int iDst,
            IImageList punkSrc,
            int iSrc,
            int uFlags);

            [PreserveSig]
            int Merge(
            int i1,
            IImageList punk2,
            int i2,
            int dx,
            int dy,
            ref Guid riid,
            ref IntPtr ppv);

            [PreserveSig]
            int Clone(
            ref Guid riid,
            ref IntPtr ppv);

            [PreserveSig]
            int GetImageRect(
            int i,
            ref RectangleI prc);

            [PreserveSig]
            int GetIconSize(
            ref int cx,
            ref int cy);

            [PreserveSig]
            int SetIconSize(
            int cx,
            int cy);

            [PreserveSig]
            int GetImageCount(
            ref int pi);

            [PreserveSig]
            int SetImageCount(
            int uNewCount);

            [PreserveSig]
            int SetBkColor(
            int clrBk,
            ref int pclr);

            [PreserveSig]
            int GetBkColor(
            ref int pclr);

            [PreserveSig]
            int BeginDrag(
            int iTrack,
            int dxHotspot,
            int dyHotspot);

            [PreserveSig]
            int EndDrag();

            [PreserveSig]
            int DragEnter(
            IntPtr hwndLock,
            int x,
            int y);

            [PreserveSig]
            int DragLeave(
            IntPtr hwndLock);

            [PreserveSig]
            int DragMove(
            int x,
            int y);

            [PreserveSig]
            int SetDragCursorImage(
            ref IImageList punk,
            int iDrag,
            int dxHotspot,
            int dyHotspot);

            [PreserveSig]
            int DragShowNolock(
            int fShow);

            [PreserveSig]
            int GetDragImage(
            ref PointI ppt,
            ref PointI pptHotspot,
            ref Guid riid,
            ref IntPtr ppv);

            [PreserveSig]
            int GetItemFlags(
            int i,
            ref int dwFlags);

            [PreserveSig]
            int GetOverlayImage(
            int iOverlay,
            ref int piIndex);
        };
        [Flags]
        public enum AnimateWindowFlags
        {
            AW_HOR_POSITIVE = 0x1,
            AW_HOR_NEGATIVE = 0x2,
            AW_VER_POSITIVE = 0x4,
            AW_VER_NEGATIVE = 0x8,
            AW_CENTER = 0x10,
            AW_HIDE = 0x10000,
            AW_ACTIVATE = 0x20000,
            AW_SLIDE = 0x40000,
            AW_BLEND = 0x80000
        }


        public enum ComboBoxButtonState
        {
            STATE_SYSTEM_NONE = 0,
            STATE_SYSTEM_INVISIBLE = 0x8000,
            STATE_SYSTEM_PRESSED = 0x8
        }


        public enum DPI_AWARENESS
        {
            DPI_AWARENESS_INVALID = -1,
            DPI_AWARENESS_UNAWARE = 0,
            DPI_AWARENESS_SYSTEM_AWARE = 1,
            DPI_AWARENESS_PER_MONITOR_AWARE = 2
        }


        public enum DPI_AWARENESS_CONTEXT
        {
            //DPI_AWARENESS_CONTEXT_UNAWARE = 16,
            //DPI_AWARENESS_CONTEXT_SYSTEM_AWARE = 17,
            //DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE = 18,
            //DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = 34
            DPI_AWARENESS_CONTEXT_UNAWARE = -1,
            DPI_AWARENESS_CONTEXT_SYSTEM_AWARE = -2,
            DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE = -3,
            DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4,
        }


        [Flags]
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
            ILD_BLEND = ILD_BLEND50,

            ILD_PRESERVEALPHA = 0x00001000,
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
            HTERROR = -2,
            HTTRANSPARENT = -1,
            HTNOWHERE = 0,
            HTCLIENT = 1,
            HTCAPTION = 2,

            HTSYSMENU = 3,
            HTGROWBOX = 4,
            HTMENU = 5,
            HTHSCROLL = 6,
            HTVSCROLL = 7,
            HTMINBUTTON = 8,
            HTMAXBUTTON = 9,
            HTLEFT = 10,
            HTRIGHT = 11,
            HTTOP = 12,
            HTTOPLEFT = 13,
            HTTOPRIGHT = 14,
            HTBOTTOM = 15,
            HTBOTTOMLEFT = 16,
            HTBOTTOMRIGHT = 17,
            HTBORDER = 18,
            HTOBJECT = 19,
            HTCLOSE = 20,
            HTHELP = 21,

            HTREDUCE = HTMINBUTTON,

            HTSIZE = HTGROWBOX,
            HTSIZEFIRST = HTLEFT,
            HTSIZELAST = HTBOTTOMRIGHT,
            HTZOOM = HTMAXBUTTON
        }


        public enum ImageType
        {
            Bitmap = 0,
            Icon = 1,
            Cursor = 2,
            EnhMetafile = 3
        }


        [Flags]
        public enum LoadImageFlags
        {
            DefaultColor = 0x0,
            Monochrome = 0x1,
            Color = 0x2,
            CopyReturnOriginal = 0x4,
            CopyDeleteOriginal = 0x8,
            LoadFromFile = 0x10,
            LoadTransparent = 0x20,
            DefaultSize = 0x40,
            VgaColor = 0x80,
            LoadMap3DColors = 0x1000,
            CreateDibSection = 0x2000,
            CopyFromResource = 0x4000,
            Shared = 0x8000
        }


        [Flags]
        public enum RedrawWindowFlags
        {
            RDW_NONE = 0x0,
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


        public enum RegionType
        {
            Error = 0,
            NULLREGION = 1,
            SIMPLEREGION = 2,
            COMPLEXREGION = 3
        }


        public enum SBOrientation
        {
            SB_HORZ = 0x0,
            SB_VERT = 0x1,
            SB_CTL = 0x2,
            SB_BOTH = 0x3
        }


        public enum ScrollInfoMask : uint
        {
            SIF_RANGE = 0x1,
            SIF_PAGE = 0x2,
            SIF_POS = 0x4,
            SIF_DISABLENOSCROLL = 0x8,
            SIF_TRACKPOS = 0x10,
            SIF_ALL = SIF_RANGE | SIF_PAGE | SIF_POS | SIF_TRACKPOS
        }




        //<System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage")>
        //<DllImport(User32, CharSet:=CharSet.Auto, SetLastError:=True)> _
        //Public Shared Function LoadImage(ByVal hInst As HandleRef, ByVal lpszName As IntPtr, ByVal uType As Int32, ByVal cxDesired As Int32, ByVal cyDesired As Int32, ByVal fuLoad As Int32) As IntPtr
        //End Function


        [Flags]
        public enum SHGSI
        {
            SHGSI_ICONLOCATION = 0,
            SHGSI_ICON = 0x100,
            SHGSI_SYSICONINDEX = 0x4000,
            SHGSI_LINKOVERLAY = 0x8000,
            SHGSI_SELECTED = 0x10000,
            SHGSI_LARGEICON = 0x0,
            SHGSI_SMALLICON = 0x1,
            SHGSI_SHELLICONSIZE = 0x4
        }

        [Flags]
        public enum SHGFI  {
            /// <summary>get icon</summary>
            Icon = 0x000000100,
            /// <summary>get display name</summary>
            DisplayName = 0x000000200,
            /// <summary>get type name</summary>
            TypeName = 0x000000400,
            /// <summary>get attributes</summary>
            Attributes = 0x000000800,
            /// <summary>get icon location</summary>
            IconLocation = 0x000001000,
            /// <summary>return exe type</summary>
            ExeType = 0x000002000,
            /// <summary>get system icon index</summary>
            SysIconIndex = 0x000004000,
            /// <summary>put a link overlay on icon</summary>
            LinkOverlay = 0x000008000,
            /// <summary>show icon in selected state</summary>
            Selected = 0x000010000,
            /// <summary>get only specified attributes</summary>
            Attr_Specified = 0x000020000,
            /// <summary>get large icon</summary>
            LargeIcon = 0x000000000,
            /// <summary>get small icon</summary>
            SmallIcon = 0x000000001,
            /// <summary>get open icon</summary>
            OpenIcon = 0x000000002,
            /// <summary>get shell size icon</summary>
            ShellIconSize = 0x000000004,
            /// <summary>pszPath is a pidl</summary>
            PIDL = 0x000000008,
            /// <summary>use passed dwFileAttribute</summary>
            UseFileAttributes = 0x000000010,
            /// <summary>apply the appropriate overlays</summary>
            AddOverlays = 0x000000020,
            /// <summary>Get the index of the overlay in the upper 8 bits of the iIcon</summary>
            OverlayIndex = 0x000000040,
 

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


        [Flags]
        public enum SPIF
        {
            None = 0x00,

            /// <summary>Writes the new system-wide parameter setting to the user profile.</summary>
            SPIF_UPDATEINIFILE = 0x01,

            /// <summary>Broadcasts the WM_SETTINGCHANGE message after updating the user profile.</summary>
            SPIF_SENDCHANGE = 0x02,

            /// <summary>Same as SPIF_SENDCHANGE.</summary>
            SPIF_SENDWININICHANGE = 0x02
        }


        /// <summary>
        ///     Flags used with the Windows API (User32.dll):GetSystemMetrics(SystemMetric smIndex)
        ///     This Enum and declaration signature was written by Gabriel T. Sharp
        ///     ai_productions@verizon.net or osirisgothra@hotmail.com
        ///     Obtained on pinvoke.net, please contribute your code to support the wiki!
        /// </summary>
        public enum SystemMetric : int
        {
            /// <summary>
            ///     The flags that specify how the system arranged minimized windows. For more information, see the Remarks section in
            ///     this topic.
            /// </summary>
            SM_ARRANGE = 56,

            /// <summary>
            ///     The value that specifies how the system is started:
            ///     0 Normal boot
            ///     1 Fail-safe boot
            ///     2 Fail-safe with network boot
            ///     A fail-safe boot (also called SafeBoot, Safe Mode, or Clean Boot) bypasses the user startup files.
            /// </summary>
            SM_CLEANBOOT = 67,

            /// <summary>
            ///     The number of display monitors on a desktop. For more information, see the Remarks section in this topic.
            /// </summary>
            SM_CMONITORS = 80,

            /// <summary>
            ///     The number of buttons on a mouse, or zero if no mouse is installed.
            /// </summary>
            SM_CMOUSEBUTTONS = 43,

            /// <summary>
            ///     The width of a window border, in pixels. This is equivalent to the SM_CXEDGE value for windows with the 3-D look.
            /// </summary>
            SM_CXBORDER = 5,

            /// <summary>
            ///     The width of a cursor, in pixels. The system cannot create cursors of other sizes.
            /// </summary>
            SM_CXCURSOR = 13,

            /// <summary>
            ///     This value is the same as SM_CXFIXEDFRAME.
            /// </summary>
            SM_CXDLGFRAME = 7,

            /// <summary>
            ///     The width of the rectangle around the location of a first click in a double-click sequence, in pixels. ,
            ///     The second click must occur within the rectangle that is defined by SM_CXDOUBLECLK and SM_CYDOUBLECLK for the
            ///     system
            ///     to consider the two clicks a double-click. The two clicks must also occur within a specified time.
            ///     To set the width of the double-click rectangle, call SystemParametersInfo with SPI_SETDOUBLECLKWIDTH.
            /// </summary>
            SM_CXDOUBLECLK = 36,

            /// <summary>
            ///     The number of pixels on either side of a mouse-down point that the mouse pointer can move before a drag operation
            ///     begins.
            ///     This allows the user to click and release the mouse button easily without unintentionally starting a drag
            ///     operation.
            ///     If this value is negative, it is subtracted from the left of the mouse-down point and added to the right of it.
            /// </summary>
            SM_CXDRAG = 68,

            /// <summary>
            ///     The width of a 3-D border, in pixels. This metric is the 3-D counterpart of SM_CXBORDER.
            /// </summary>
            SM_CXEDGE = 45,

            /// <summary>
            ///     The thickness of the frame around the perimeter of a window that has a caption but is not sizable, in pixels.
            ///     SM_CXFIXEDFRAME is the height of the horizontal border, and SM_CYFIXEDFRAME is the width of the vertical border.
            ///     This value is the same as SM_CXDLGFRAME.
            /// </summary>
            SM_CXFIXEDFRAME = 7,

            /// <summary>
            ///     The width of the left and right edges of the focus rectangle that the DrawFocusRectdraws.
            ///     This value is in pixels.
            ///     Windows 2000:  This value is not supported.
            /// </summary>
            SM_CXFOCUSBORDER = 83,

            /// <summary>
            ///     This value is the same as SM_CXSIZEFRAME.
            /// </summary>
            SM_CXFRAME = 32,

            /// <summary>
            ///     The width of the client area for a full-screen window on the primary display monitor, in pixels.
            ///     To get the coordinates of the portion of the screen that is not obscured by the system taskbar or by application
            ///     desktop toolbars,
            ///     call the SystemParametersInfofunction with the SPI_GETWORKAREA value.
            /// </summary>
            SM_CXFULLSCREEN = 16,

            /// <summary>
            ///     The width of the arrow bitmap on a horizontal scroll bar, in pixels.
            /// </summary>
            SM_CXHSCROLL = 21,

            /// <summary>
            ///     The width of the thumb box in a horizontal scroll bar, in pixels.
            /// </summary>
            SM_CXHTHUMB = 10,

            /// <summary>
            ///     The default width of an icon, in pixels. The LoadIcon function can load only icons with the dimensions
            ///     that SM_CXICON and SM_CYICON specifies.
            /// </summary>
            SM_CXICON = 11,

            /// <summary>
            ///     The width of a grid cell for items in large icon view, in pixels. Each item fits into a rectangle of size
            ///     SM_CXICONSPACING by SM_CYICONSPACING when arranged. This value is always greater than or equal to SM_CXICON.
            /// </summary>
            SM_CXICONSPACING = 38,

            /// <summary>
            ///     The default width, in pixels, of a maximized top-level window on the primary display monitor.
            /// </summary>
            SM_CXMAXIMIZED = 61,

            /// <summary>
            ///     The default maximum width of a window that has a caption and sizing borders, in pixels.
            ///     This metric refers to the entire desktop. The user cannot drag the window frame to a size larger than these
            ///     dimensions.
            ///     A window can override this value by processing the WM_GETMINMAXINFO message.
            /// </summary>
            SM_CXMAXTRACK = 59,

            /// <summary>
            ///     The width of the default menu check-mark bitmap, in pixels.
            /// </summary>
            SM_CXMENUCHECK = 71,

            /// <summary>
            ///     The width of menu bar buttons, such as the child window close button that is used in the multiple document
            ///     interface, in pixels.
            /// </summary>
            SM_CXMENUSIZE = 54,

            /// <summary>
            ///     The minimum width of a window, in pixels.
            /// </summary>
            SM_CXMIN = 28,

            /// <summary>
            ///     The width of a minimized window, in pixels.
            /// </summary>
            SM_CXMINIMIZED = 57,

            /// <summary>
            ///     The width of a grid cell for a minimized window, in pixels. Each minimized window fits into a rectangle this size
            ///     when arranged.
            ///     This value is always greater than or equal to SM_CXMINIMIZED.
            /// </summary>
            SM_CXMINSPACING = 47,

            /// <summary>
            ///     The minimum tracking width of a window, in pixels. The user cannot drag the window frame to a size smaller than
            ///     these dimensions.
            ///     A window can override this value by processing the WM_GETMINMAXINFO message.
            /// </summary>
            SM_CXMINTRACK = 34,

            /// <summary>
            ///     The amount of border padding for captioned windows, in pixels. Windows XP/2000:  This value is not supported.
            /// </summary>
            SM_CXPADDEDBORDER = 92,

            /// <summary>
            ///     The width of the screen of the primary display monitor, in pixels. This is the same value obtained by calling
            ///     GetDeviceCaps as follows: GetDeviceCaps( hdcPrimaryMonitor, HORZRES).
            /// </summary>
            SM_CXSCREEN = 0,

            /// <summary>
            ///     The width of a button in a window caption or title bar, in pixels.
            /// </summary>
            SM_CXSIZE = 30,

            /// <summary>
            ///     The thickness of the sizing border around the perimeter of a window that can be resized, in pixels.
            ///     SM_CXSIZEFRAME is the width of the horizontal border, and SM_CYSIZEFRAME is the height of the vertical border.
            ///     This value is the same as SM_CXFRAME.
            /// </summary>
            SM_CXSIZEFRAME = 32,

            /// <summary>
            ///     The recommended width of a small icon, in pixels. Small icons typically appear in window captions and in small icon
            ///     view.
            /// </summary>
            SM_CXSMICON = 49,

            /// <summary>
            ///     The width of small caption buttons, in pixels.
            /// </summary>
            SM_CXSMSIZE = 52,

            /// <summary>
            ///     The width of the virtual screen, in pixels. The virtual screen is the bounding rectangle of all display monitors.
            ///     The SM_XVIRTUALSCREEN metric is the coordinates for the left side of the virtual screen.
            /// </summary>
            SM_CXVIRTUALSCREEN = 78,

            /// <summary>
            ///     The width of a vertical scroll bar, in pixels.
            /// </summary>
            SM_CXVSCROLL = 2,

            /// <summary>
            ///     The height of a window border, in pixels. This is equivalent to the SM_CYEDGE value for windows with the 3-D look.
            /// </summary>
            SM_CYBORDER = 6,

            /// <summary>
            ///     The height of a caption area, in pixels.
            /// </summary>
            SM_CYCAPTION = 4,

            /// <summary>
            ///     The height of a cursor, in pixels. The system cannot create cursors of other sizes.
            /// </summary>
            SM_CYCURSOR = 14,

            /// <summary>
            ///     This value is the same as SM_CYFIXEDFRAME.
            /// </summary>
            SM_CYDLGFRAME = 8,

            /// <summary>
            ///     The height of the rectangle around the location of a first click in a double-click sequence, in pixels.
            ///     The second click must occur within the rectangle defined by SM_CXDOUBLECLK and SM_CYDOUBLECLK for the system to
            ///     consider
            ///     the two clicks a double-click. The two clicks must also occur within a specified time. To set the height of the
            ///     double-click
            ///     rectangle, call SystemParametersInfo with SPI_SETDOUBLECLKHEIGHT.
            /// </summary>
            SM_CYDOUBLECLK = 37,

            /// <summary>
            ///     The number of pixels above and below a mouse-down point that the mouse pointer can move before a drag operation
            ///     begins.
            ///     This allows the user to click and release the mouse button easily without unintentionally starting a drag
            ///     operation.
            ///     If this value is negative, it is subtracted from above the mouse-down point and added below it.
            /// </summary>
            SM_CYDRAG = 69,

            /// <summary>
            ///     The height of a 3-D border, in pixels. This is the 3-D counterpart of SM_CYBORDER.
            /// </summary>
            SM_CYEDGE = 46,

            /// <summary>
            ///     The thickness of the frame around the perimeter of a window that has a caption but is not sizable, in pixels.
            ///     SM_CXFIXEDFRAME is the height of the horizontal border, and SM_CYFIXEDFRAME is the width of the vertical border.
            ///     This value is the same as SM_CYDLGFRAME.
            /// </summary>
            SM_CYFIXEDFRAME = 8,

            /// <summary>
            ///     The height of the top and bottom edges of the focus rectangle drawn byDrawFocusRect.
            ///     This value is in pixels.
            ///     Windows 2000:  This value is not supported.
            /// </summary>
            SM_CYFOCUSBORDER = 84,

            /// <summary>
            ///     This value is the same as SM_CYSIZEFRAME.
            /// </summary>
            SM_CYFRAME = 33,

            /// <summary>
            ///     The height of the client area for a full-screen window on the primary display monitor, in pixels.
            ///     To get the coordinates of the portion of the screen not obscured by the system taskbar or by application desktop
            ///     toolbars,
            ///     call the SystemParametersInfo function with the SPI_GETWORKAREA value.
            /// </summary>
            SM_CYFULLSCREEN = 17,

            /// <summary>
            ///     The height of a horizontal scroll bar, in pixels.
            /// </summary>
            SM_CYHSCROLL = 3,

            /// <summary>
            ///     The default height of an icon, in pixels. The LoadIcon function can load only icons with the dimensions SM_CXICON
            ///     and SM_CYICON.
            /// </summary>
            SM_CYICON = 12,

            /// <summary>
            ///     The height of a grid cell for items in large icon view, in pixels. Each item fits into a rectangle of size
            ///     SM_CXICONSPACING by SM_CYICONSPACING when arranged. This value is always greater than or equal to SM_CYICON.
            /// </summary>
            SM_CYICONSPACING = 39,

            /// <summary>
            ///     For double byte character set versions of the system, this is the height of the Kanji window at the bottom of the
            ///     screen, in pixels.
            /// </summary>
            SM_CYKANJIWINDOW = 18,

            /// <summary>
            ///     The default height, in pixels, of a maximized top-level window on the primary display monitor.
            /// </summary>
            SM_CYMAXIMIZED = 62,

            /// <summary>
            ///     The default maximum height of a window that has a caption and sizing borders, in pixels. This metric refers to the
            ///     entire desktop.
            ///     The user cannot drag the window frame to a size larger than these dimensions. A window can override this value by
            ///     processing
            ///     the WM_GETMINMAXINFO message.
            /// </summary>
            SM_CYMAXTRACK = 60,

            /// <summary>
            ///     The height of a single-line menu bar, in pixels.
            /// </summary>
            SM_CYMENU = 15,

            /// <summary>
            ///     The height of the default menu check-mark bitmap, in pixels.
            /// </summary>
            SM_CYMENUCHECK = 72,

            /// <summary>
            ///     The height of menu bar buttons, such as the child window close button that is used in the multiple document
            ///     interface, in pixels.
            /// </summary>
            SM_CYMENUSIZE = 55,

            /// <summary>
            ///     The minimum height of a window, in pixels.
            /// </summary>
            SM_CYMIN = 29,

            /// <summary>
            ///     The height of a minimized window, in pixels.
            /// </summary>
            SM_CYMINIMIZED = 58,

            /// <summary>
            ///     The height of a grid cell for a minimized window, in pixels. Each minimized window fits into a rectangle this size
            ///     when arranged.
            ///     This value is always greater than or equal to SM_CYMINIMIZED.
            /// </summary>
            SM_CYMINSPACING = 48,

            /// <summary>
            ///     The minimum tracking height of a window, in pixels. The user cannot drag the window frame to a size smaller than
            ///     these dimensions.
            ///     A window can override this value by processing the WM_GETMINMAXINFO message.
            /// </summary>
            SM_CYMINTRACK = 35,

            /// <summary>
            ///     The height of the screen of the primary display monitor, in pixels. This is the same value obtained by calling
            ///     GetDeviceCaps as follows: GetDeviceCaps( hdcPrimaryMonitor, VERTRES).
            /// </summary>
            SM_CYSCREEN = 1,

            /// <summary>
            ///     The height of a button in a window caption or title bar, in pixels.
            /// </summary>
            SM_CYSIZE = 31,

            /// <summary>
            ///     The thickness of the sizing border around the perimeter of a window that can be resized, in pixels.
            ///     SM_CXSIZEFRAME is the width of the horizontal border, and SM_CYSIZEFRAME is the height of the vertical border.
            ///     This value is the same as SM_CYFRAME.
            /// </summary>
            SM_CYSIZEFRAME = 33,

            /// <summary>
            ///     The height of a small caption, in pixels.
            /// </summary>
            SM_CYSMCAPTION = 51,

            /// <summary>
            ///     The recommended height of a small icon, in pixels. Small icons typically appear in window captions and in small
            ///     icon view.
            /// </summary>
            SM_CYSMICON = 50,

            /// <summary>
            ///     The height of small caption buttons, in pixels.
            /// </summary>
            SM_CYSMSIZE = 53,

            /// <summary>
            ///     The height of the virtual screen, in pixels. The virtual screen is the bounding rectangle of all display monitors.
            ///     The SM_YVIRTUALSCREEN metric is the coordinates for the top of the virtual screen.
            /// </summary>
            SM_CYVIRTUALSCREEN = 79,

            /// <summary>
            ///     The height of the arrow bitmap on a vertical scroll bar, in pixels.
            /// </summary>
            SM_CYVSCROLL = 20,

            /// <summary>
            ///     The height of the thumb box in a vertical scroll bar, in pixels.
            /// </summary>
            SM_CYVTHUMB = 9,

            /// <summary>
            ///     Nonzero if User32.dll supports DBCS; otherwise, 0.
            /// </summary>
            SM_DBCSENABLED = 42,

            /// <summary>
            ///     Nonzero if the debug version of User.exe is installed; otherwise, 0.
            /// </summary>
            SM_DEBUG = 22,

            /// <summary>
            ///     Nonzero if the current operating system is Windows 7 or Windows Server 2008 R2 and the Tablet PC Input
            ///     service is started; otherwise, 0. The return value is a bitmask that specifies the type of digitizer input
            ///     supported by the device.
            ///     For more information, see Remarks.
            ///     Windows Server 2008, Windows Vista, and Windows XP/2000:  This value is not supported.
            /// </summary>
            SM_DIGITIZER = 94,

            /// <summary>
            ///     Nonzero if Input Method Manager/Input Method Editor features are enabled; otherwise, 0.
            ///     SM_IMMENABLED indicates whether the system is ready to use a Unicode-based IME on a Unicode application.
            ///     To ensure that a language-dependent IME works, check SM_DBCSENABLED and the system ANSI code page.
            ///     Otherwise the ANSI-to-Unicode conversion may not be performed correctly, or some components like fonts
            ///     or registry settings may not be present.
            /// </summary>
            SM_IMMENABLED = 82,

            /// <summary>
            ///     Nonzero if there are digitizers in the system; otherwise, 0. SM_MAXIMUMTOUCHES returns the aggregate maximum of the
            ///     maximum number of contacts supported by every digitizer in the system. If the system has only single-touch
            ///     digitizers,
            ///     the return value is 1. If the system has multi-touch digitizers, the return value is the number of simultaneous
            ///     contacts
            ///     the hardware can provide. Windows Server 2008, Windows Vista, and Windows XP/2000:  This value is not supported.
            /// </summary>
            SM_MAXIMUMTOUCHES = 95,

            /// <summary>
            ///     Nonzero if the current operating system is the Windows XP, Media Center Edition, 0 if not.
            /// </summary>
            SM_MEDIACENTER = 87,

            /// <summary>
            ///     Nonzero if drop-down menus are right-aligned with the corresponding menu-bar item; 0 if the menus are left-aligned.
            /// </summary>
            SM_MENUDROPALIGNMENT = 40,

            /// <summary>
            ///     Nonzero if the system is enabled for Hebrew and Arabic languages, 0 if not.
            /// </summary>
            SM_MIDEASTENABLED = 74,

            /// <summary>
            ///     Nonzero if a mouse is installed; otherwise, 0. This value is rarely zero, because of support for virtual mice and
            ///     because
            ///     some systems detect the presence of the port instead of the presence of a mouse.
            /// </summary>
            SM_MOUSEPRESENT = 19,

            /// <summary>
            ///     Nonzero if a mouse with a horizontal scroll wheel is installed; otherwise 0.
            /// </summary>
            SM_MOUSEHORIZONTALWHEELPRESENT = 91,

            /// <summary>
            ///     Nonzero if a mouse with a vertical scroll wheel is installed; otherwise 0.
            /// </summary>
            SM_MOUSEWHEELPRESENT = 75,

            /// <summary>
            ///     The least significant bit is set if a network is present; otherwise, it is cleared. The other bits are reserved for
            ///     future use.
            /// </summary>
            SM_NETWORK = 63,

            /// <summary>
            ///     Nonzero if the Microsoft Windows for Pen computing extensions are installed; zero otherwise.
            /// </summary>
            SM_PENWINDOWS = 41,

            /// <summary>
            ///     This system metric is used in a Terminal Services environment to determine if the current Terminal Server session
            ///     is
            ///     being remotely controlled. Its value is nonzero if the current session is remotely controlled; otherwise, 0.
            ///     You can use terminal services management tools such as Terminal Services Manager (tsadmin.msc) and shadow.exe to
            ///     control a remote session. When a session is being remotely controlled, another user can view the contents of that
            ///     session
            ///     and potentially interact with it.
            /// </summary>
            SM_REMOTECONTROL = 0x2001,

            /// <summary>
            ///     This system metric is used in a Terminal Services environment. If the calling process is associated with a Terminal
            ///     Services
            ///     client session, the return value is nonzero. If the calling process is associated with the Terminal Services
            ///     console session,
            ///     the return value is 0.
            ///     Windows Server 2003 and Windows XP:  The console session is not necessarily the physical console.
            ///     For more information, seeWTSGetActiveConsoleSessionId.
            /// </summary>
            SM_REMOTESESSION = 0x1000,

            /// <summary>
            ///     Nonzero if all the display monitors have the same color format, otherwise, 0. Two displays can have the same bit
            ///     depth,
            ///     but different color formats. For example, the red, green, and blue pixels can be encoded with different numbers of
            ///     bits,
            ///     or those bits can be located in different places in a pixel color value.
            /// </summary>
            SM_SAMEDISPLAYFORMAT = 81,

            /// <summary>
            ///     This system metric should be ignored; it always returns 0.
            /// </summary>
            SM_SECURE = 44,

            /// <summary>
            ///     The build number if the system is Windows Server 2003 R2; otherwise, 0.
            /// </summary>
            SM_SERVERR2 = 89,

            /// <summary>
            ///     Nonzero if the user requires an application to present information visually in situations where it would otherwise
            ///     present
            ///     the information only in audible form; otherwise, 0.
            /// </summary>
            SM_SHOWSOUNDS = 70,

            /// <summary>
            ///     Nonzero if the current session is shutting down; otherwise, 0. Windows 2000:  This value is not supported.
            /// </summary>
            SM_SHUTTINGDOWN = 0x2000,

            /// <summary>
            ///     Nonzero if the computer has a low-end (slow) processor; otherwise, 0.
            /// </summary>
            SM_SLOWMACHINE = 73,

            /// <summary>
            ///     Nonzero if the current operating system is Windows 7 Starter Edition, Windows Vista Starter, or Windows XP Starter
            ///     Edition; otherwise, 0.
            /// </summary>
            SM_STARTER = 88,

            /// <summary>
            ///     Nonzero if the meanings of the left and right mouse buttons are swapped; otherwise, 0.
            /// </summary>
            SM_SWAPBUTTON = 23,

            /// <summary>
            ///     Nonzero if the current operating system is the Windows XP Tablet PC edition or if the current operating system is
            ///     Windows Vista
            ///     or Windows 7 and the Tablet PC Input service is started; otherwise, 0. The SM_DIGITIZER setting indicates the type
            ///     of digitizer
            ///     input supported by a device running Windows 7 or Windows Server 2008 R2. For more information, see Remarks.
            /// </summary>
            SM_TABLETPC = 86,

            /// <summary>
            ///     The coordinates for the left side of the virtual screen. The virtual screen is the bounding rectangle of all
            ///     display monitors.
            ///     The SM_CXVIRTUALSCREEN metric is the width of the virtual screen.
            /// </summary>
            SM_XVIRTUALSCREEN = 76,

            /// <summary>
            ///     The coordinates for the top of the virtual screen. The virtual screen is the bounding rectangle of all display
            ///     monitors.
            ///     The SM_CYVIRTUALSCREEN metric is the height of the virtual screen.
            /// </summary>
            SM_YVIRTUALSCREEN = 77,
        }


        public enum SystemParameters
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

            SPI_GETCLIENTAREAANIMATION = 0x1042,

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





        
        public const int S_OK = 0;
        public const int SC_CLOSE = 0xF060;

        public const int ERROR_CLIPBOARD_CANT_OPEN = unchecked((int) 0x800401D0);
        public const int ERROR_CLIPBOARD_CANT_EMPTY = unchecked((int) 0x800401D1);
        public const int ERROR_CLIPBOARD_CANT_SET = unchecked((int) 0x800401D2);
        public const int ERROR_CLIPBOARD_BAD_DATA = unchecked((int) 0x800401D3);
        public const int ERROR_CLIPBOARD_CANT_CLOSE = unchecked((int) 0x800401D4);

        public const int GCL_HBRBACKGROUND = -10;
        public const int COLOR_WINDOW = 5;

        /// <summary>
        ///     Sets the elevation required state for a specified button or
        ///     command link to display an elevated icon.
        /// </summary>
        public const int BCM_SETSHIELD = 0x160C;

        public const int MA_ACTIVATE = 0x0001;
        public const int MA_ACTIVATEANDEAT = 0x0002;
        public const int MA_NOACTIVATE = 0x0003;
        public const int MA_NOACTIVATEANDEAT = 0x0004;


        public const int CB_GETCOMBOBOXINFO = 0x164;
        public const int LB_ITEMFROMPOINT = 0x1A9;
        public const int LB_GETITEMRECT = 0x198;

        public const int CBN_SELCHANGE = 0x1;


        public const int WVR_HREDRAW = 0x100;
        public const int WVR_VREDRAW = 0x200;
        public const int WVR_REDRAW = WVR_HREDRAW | WVR_VREDRAW;
        public const int WVR_VALIDRECTS = 0x400;
        public const int PRF_NONCLIENT = 0x2;

        public const int DCX_INTERSECTRGN = 0x80;
        public const int DCX_WINDOW = 0x1;
        public const int DCX_CACHE = 0x2;
        public const int DCX_LOCKWINDOWUPDATE = 1024;
        public const int DCX_USESTYLE = 0x00010000;//UnDocumented!
 
  
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
        public const int GWL_HWNDPARENT = -8;

        public const int ULW_COLORKEY = 0x1;
        public const int ULW_ALPHA = 0x2;
        public const int ULW_OPAQUE = 0x4;

        public const int HWND_BOTTOM = 1;
        public const int HWND_DESKTOP = 0;
        public const int HWND_NOTOPMOST = -2;
        public const int HWND_TOP = 0;
        public const int HWND_TOPMOST = -1;

        public const int GCW_ATOM = -32;

        public const int MAX_PATH = 260;

        public const int WS_EX_LAYERED = 0x80000;

        public const int OBJID_HSCROLL = unchecked((int) 0xFFFFFFFA);
        public const int OBJID_VSCROLL = unchecked((int) 0xFFFFFFFB);
        public const int OBJID_CLIENT = unchecked((int) 0xFFFFFFFC);

        public const int SB_CTL = 2;
        public const int SB_ENDSCROLL = 8;

        public const int LB_GETCURSEL = 0x188;

        public const int LOGPIXELSX = 0x58;
        public const int LOGPIXELSY = 90;


        internal const string TOOLTIPS_CLASS = "tooltips_class32";
        internal const int TTS_NOPREFIX = 0x2;
        internal const int TTS_USEVISUALSTYLE = 0x100;
        internal const int TTS_ALWAYSTIP = 0x1;
        internal const int TTS_NOANIMATE = 0x10;
        internal const int TTS_NOFADE = 0x20;
        internal const int TTS_BALLOON = 0x40;


        internal const int TTF_IDISHWND = 0x1;
        internal const int TTF_TRACK = 0x20;
        internal const int TTF_ABSOLUTE = 0x80;
        internal const int TTF_TRANSPARENT = 0x100;
        internal const int TTF_SUBCLASS = 0x10;
        internal const int TTF_PARSELINKS = 0x1000;

        internal const int TTM_UPDATE = WindowMessages.WM_USER + 29;

        internal const int TTM_ACTIVATE = WindowMessages.WM_USER + 1;
        internal const int TTM_ADDTOOLA = WindowMessages.WM_USER + 4;
        internal const int TTM_DELTOOL = WindowMessages.WM_USER + +51;
        internal const int TTM_TRACKACTIVATE = WindowMessages.WM_USER + 17;
        internal const int TTM_TRACKPOSITION = WindowMessages.WM_USER + 18;
        internal const int TTM_SETMAXTIPWIDTH = WindowMessages.WM_USER + 24;
        internal const int TTM_SETTITLE = WindowMessages.WM_USER + 33;
        internal const int TTM_ADDTOOL = WindowMessages.WM_USER + 50;
        internal const int TTM_UPDATETIPTEXT = WindowMessages.WM_USER + 57;
        internal const int TTM_GETBUBBLESIZE = WindowMessages.WM_USER + 30;
        internal const int TTM_POP = WindowMessages.WM_USER + 28;
        internal const int TTM_POPUP = WindowMessages.WM_USER + 27;

        internal const int TTN_SHOW = -521;
        internal const int TTN_POP = -522;

        private const int TTTOOLINFO_V2_SIZE = 44;


        public const long ASSOCSTR_EXECUTABLE = 2;
        public const long ASSOCF_IGNOREUNKNOWN = 0x400;


        public const int CWP_ALL = 0x0;
        public const int CWP_SKIPDISABLED = 0x2;
        public const int CWP_SKIPINVISIBLE = 0x1;
        public const int CWP_SKIPTRANSPARENT = 0x4;

        public const int MM_ISOTROPIC = 7;
        public const int MM_ANISOTROPIC = 8;
        public const int MM_TEXT = 1;
        public const int MM_LOENGLISH = 4;
        public const int MM_HIMETRIC = 3;

        public const uint FILE_ATTRIBUTE_NORMAL = 0x80;
        public const uint FILE_ATTRIBUTE_DIRECTORY = 0x10;
        
        // BlendOp:
        public const int AC_SRC_OVER = 0x0;

        // AlphaFormat:
        public const int AC_SRC_ALPHA = 0x1;


        public const int OBJID_NATIVEOM = -16;

        public static readonly HandleRef NullHandleRef = new HandleRef(wrapper: null, handle: IntPtr.Zero);

        public static int WM_LJ_OWNERDPICHANGED;
        


        private static int _currentProcessId;
	
		static NativeMethods()
        {
            WM_LJ_OWNERDPICHANGED = RegisterWindowMessage(nameof(WM_LJ_OWNERDPICHANGED));
        }

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
        [DllImport(Shcore)]
        private static extern uint GetDpiForMonitor(IntPtr hmonitor, uint dpiType, out int dpiX, out int dpiY);

        [DllImport(Shcore, EntryPoint = "GetScaleFactorForMonitor")]
        private static extern int GetScaleFactorForMonitorAPI(IntPtr hMon, out int pScale);

        public static int GetDpiFromRect(RectangleI r)
        {
            var hMon = MonitorFromRectAPI(r, MONITOR_DEFAULTTONEAREST);
            var sf = GetScaleFactorForMonitor(hMon);
            var pDpi = (int)(96F / sf);
            return pDpi;
        }

        public static float GetScaleFactorFromPoint(PointI pt)
        {
            var hMon = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
            return GetScaleFactorForMonitor(hMon);
        }

        public static void GetDpiFromPoint(PointI pt,out int dpiX, out int dpiY)
        {
            var hMon = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
            GetDpiForMonitor(hMon, 0,out dpiX, out dpiY);

        }

        public static float GetPrimaryDpi()
        {
            var hDc = GetDC(new HandleRef(wrapper: null, handle: IntPtr.Zero));
            var logx = GetDeviceCaps(hDc, LOGPIXELSX);
            ReleaseDC(new HandleRef(wrapper: null, handle: IntPtr.Zero), hDc);
            return logx;
        }

 
        #region      GetKeyState 

        [DllImport(User32, SetLastError = false)]
        public static extern int GetAsyncKeyState(int vKey);

        #endregion // GetKeyState

        public static float GetScaleFactorForMonitor(IntPtr hMon)
        {
            if (!OSVersionHelper.IsWindows8Point1OrGreater())
                return GetPrimaryScaleFactor();

            if (GetScaleFactorForMonitorAPI(hMon, out var pScale) != 0)
                return GetPrimaryScaleFactor();

            return pScale ;

            float GetPrimaryScaleFactor()
            {
                var hDc = GetDC(new HandleRef(wrapper: null, handle: IntPtr.Zero));
                var logx = GetDeviceCaps(hDc, LOGPIXELSX);
                ReleaseDC(new HandleRef(wrapper: null, handle: IntPtr.Zero), hDc);
                return (float)logx / 96F;
            }
        }


        public const int MONITOR_DEFAULTTONULL = 0;
        public const int MONITOR_DEFAULTTOPRIMARY = 1;
        public const int MONITOR_DEFAULTTONEAREST = 2;

        [DllImport(User32, ExactSpelling = true)]
        public static extern IntPtr MonitorFromPoint(PointI pt, int flags);

        [DllImport(User32, ExactSpelling = true, EntryPoint = "MonitorFromRect")]
        private static extern IntPtr MonitorFromRectAPI(RectangleI Rect, int flags);

        public static IntPtr MonitorFromRect(RectangleI r, int flags)
        {
            return MonitorFromRectAPI(r, flags);
        }

        [DllImport(User32, ExactSpelling = true)]
        public static extern IntPtr MonitorFromWindow(HandleRef handle, int flags);

        [DllImport(User32, CharSet = CharSet.Auto)]
        public static extern bool GetMonitorInfo(HandleRef hmonitor, [In, Out] MONITORINFOEX info);

        [DllImport(User32, CharSet = CharSet.Auto)]
        public static extern int RegisterWindowMessage(string msg);


        [DllImport(User32, EntryPoint = "SetThreadDpiAwarenessContext")]
        private static extern DpiAwarenessContextHandle SetThreadDpiAwarenessContextApi(DpiAwarenessContextHandle dpiContext);

        //[DebuggerStepThrough]
        public static DpiAwarenessContextHandle SetThreadDpiAwarenessContext(DpiAwarenessContextHandle dpiContext)
        {
            if (!OSVersionHelper.IsWindows10AnniversaryAdditionOrGreater())
                return default(DpiAwarenessContextHandle);

            return SetThreadDpiAwarenessContextApi(dpiContext);
        }


        [DllImport(User32, EntryPoint = "GetAwarenessFromDpiAwarenessContext")]
        private static extern DPI_AWARENESS GetAwarenessFromDpiAwarenessContextApi(DpiAwarenessContextHandle context);

        [DebuggerStepThrough]
        public static DPI_AWARENESS GetAwarenessFromDpiAwarenessContext(DpiAwarenessContextHandle context)
        {
            if (!OSVersionHelper.IsWindows10AnniversaryAdditionOrGreater())
                return DPI_AWARENESS.DPI_AWARENESS_INVALID;

            return GetAwarenessFromDpiAwarenessContextApi(context);
        }


        [DllImport(User32, EntryPoint = "GetThreadDpiAwarenessContext")]
        private static extern DpiAwarenessContextHandle GetThreadDpiAwarenessContextApi();

        [DebuggerStepThrough]
        public static DpiAwarenessContextHandle GetThreadDpiAwarenessContext()
        {
            if (!OSVersionHelper.IsWindows10AnniversaryAdditionOrGreater())
                return default(DpiAwarenessContextHandle);

            return GetThreadDpiAwarenessContextApi();
        }

        
        [DllImport(User32, EntryPoint = "EnableNonClientDpiScaling")]
        private static extern bool EnableNonClientDpiScalingApi(HandleRef handle);

        [DebuggerStepThrough]
        public static bool EnableNonClientDpiScaling(HandleRef handle)
        {
            if (!OSVersionHelper.IsWindows10AnniversaryAdditionOrGreater())
                return false;

            return EnableNonClientDpiScalingApi(handle);
        }


        [DllImport(User32, EntryPoint = "GetWindowDpiAwarenessContext")]
        private static extern DpiAwarenessContextHandle GetWindowDpiAwarenessContextApi(HandleRef handle);

        [DebuggerStepThrough]
        public static DpiAwarenessContextHandle GetWindowDpiAwarenessContext(HandleRef handle)
        {
            if (!OSVersionHelper.IsWindows10AnniversaryAdditionOrGreater())
                return default(DpiAwarenessContextHandle);

            return GetWindowDpiAwarenessContextApi(handle);
        }

        [DllImport(User32, EntryPoint = "AreDpiAwarenessContextsEqual")]
        private static extern bool AreDpiAwarenessContextsEqualApi(DpiAwarenessContextHandle context1, DpiAwarenessContextHandle context2);
 
        [DebuggerStepThrough]
        public static bool AreDpiAwarenessContextsEqual(DpiAwarenessContextHandle context1, DpiAwarenessContextHandle context2)
        {
            if (!OSVersionHelper.IsWindows10AnniversaryAdditionOrGreater())
                return context1.Handle == context1.Handle;

            return AreDpiAwarenessContextsEqualApi(context1, context2);
        }

        [DllImport(User32, EntryPoint = "GetDpiForWindow")]
        private static extern int GetDpiForWindowApi(IntPtr handle);

        [DebuggerStepThrough]
        public static int GetDpiForWindow(IntPtr handle)
        {
            if (!OSVersionHelper.IsWindows10AnniversaryAdditionOrGreater())
            {
                var hDc = GetDC(new HandleRef(wrapper: null, handle: IntPtr.Zero));
                var scaleFactor = GetDeviceCaps(hDc, LOGPIXELSX);
                ReleaseDC(new HandleRef(wrapper: null, handle: IntPtr.Zero), hDc);
                return scaleFactor;
            }

            return GetDpiForWindowApi(handle);
        }


        [DllImport(User32, EntryPoint = "GetSystemMetricsForDpi")]
        private static extern int GetSystemMetricsForDpiApi(int index, int dp);

        [DllImport(User32)]
        public static extern int GetSystemMetrics(int index);

        [DebuggerStepThrough]
        public static int GetSystemMetricsForDpi(SystemMetric index, int dpi)
        {
            if (!OSVersionHelper.IsWindows10AnniversaryAdditionOrGreater())
                return GetSystemMetrics((int) index);

            return GetSystemMetricsForDpiApi((int) index, dpi);
        }


        [DllImport(Gdi32, CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        private static extern bool OffsetViewportOrgEx(HandleRef hDC, int nXOffset, int nYOffset, out PointI point);

        [DllImport(Gdi32, CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        public static extern int CombineRgn(HandleRef hRgnDest, HandleRef hRgnSrc1, HandleRef hRgnSrc2, int combineMode);

        [DllImport(Gdi32, CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        public static extern int GetClipRgn(HandleRef hDC, HandleRef hRgn);

        [DllImport(Gdi32, CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        public static extern int SelectClipRgn(HandleRef hDC, HandleRef hRgn);


        //[DllImport(Uxtheme, CharSet = CharSet.Auto)]
        //public static extern int GetThemePartSize(HandleRef hTheme, HandleRef hdc, int iPartId, int iStateId, [In] COMRECT prc, ThemeSizeType eSize, [Out] COMSIZE psz);

        //[DllImport(Uxtheme, CharSet = CharSet.Auto)]
        //public static extern IntPtr OpenThemeData(HandleRef hwnd, [MarshalAs(UnmanagedType.LPWStr)] string pszClassList);


        //[DllImport(Uxtheme, CharSet = CharSet.Auto)]
        //public static extern int CloseThemeData(HandleRef hTheme);


        //[DllImport(Uxtheme, CharSet = CharSet.Auto)]
        //private static extern IntPtr OpenThemeDataForDpi(HandleRef hwnd, [MarshalAs(UnmanagedType.LPWStr)] string pszClassList, int dpi);

 
        //public static Size GetVisualStylePartSize(IntPtr hWnd, int deviceDpi, IDeviceContext dc, VisualStyleRenderer renderer)
        //{
        //    if (deviceDpi == (int) UIHelper.PrimaryDpi.X || !OSVersionHelper.IsWindows10CreatorsAdditionOrGreater())
        //        return renderer.GetPartSize(dc, ThemeSizeType.True);

        //    var handleRef = new HandleRef(null, hWnd);
        //    var size = new NativeMethods.COMSIZE();
        //    var handle = OpenThemeDataForDpi(handleRef, renderer.Class, deviceDpi);
        //    if (handle != IntPtr.Zero)
        //    {
        //        var hDc = new HandleRef(dc, dc.GetHdc());
        //        var themeHandle = new HandleRef(null, handle);

        //        //ThemeSizeType.Draw does not work with TaskDialog cheveron parts
        //        GetThemePartSize(themeHandle, hDc, renderer.Part, renderer.State,
        //                         null, ThemeSizeType.True, size);
                
        //        dc.ReleaseHdc();

        //        CloseThemeData(themeHandle);
        //    }

        //    return size.ToSize;
        //}

        [DllImport(User32, EntryPoint = "SystemParametersInfoForDpi")]
        private static extern int SystemParametersInfoForDpiApi(int index, int dp);


        [DebuggerStepThrough]
        public static int SystemParametersInfo(SystemParameters index, int dpi)
        {
            if (!OSVersionHelper.IsWindows10AnniversaryAdditionOrGreater())
                return GetSystemMetrics((int) index);

            return GetSystemMetricsForDpiApi((int) index, dpi);
        }


        [DllImport(Gdi32)]
        public static extern bool GetViewportOrgEx(HandleRef hdc, out PointI lpPoint);

        //[DllImport(Kernel32)]
        //public static extern uint GetTickCount();

        //[DllImport(Uxtheme, CharSet = CharSet.Auto)]
        //private static extern int DrawThemeBackground(HandleRef hTheme, HandleRef hdc, int partId, int stateId, [In] NativeMethods.COMRECT pRect, [In] NativeMethods.COMRECT pClipRect);


        //public static void DrawThemeBackground(VisualStyleRenderer renderer, Graphics dc, Rectangle bounds, IntPtr hWnd, int dpi)
        //{
        //    if ((bounds.Width == 0) || (bounds.Height == 0))
        //        return;

        //    if ( dpi == 0 || hWnd == IntPtr.Zero || !OSVersionHelper.IsWindows10CreatorsAdditionOrGreater())
        //    {
        //        renderer.DrawBackground(dc, bounds);
        //        return;
        //    }

        //    var matrix = dc.Transform;
        //    HandleRef hdc = default(HandleRef);
        //    HandleRef hTheme = default(HandleRef);

        //    try
        //    {
        //        bounds.X += (int) (matrix.OffsetX);
        //        bounds.Y += (int) (matrix.OffsetY);

        //        hTheme = new HandleRef(renderer, OpenThemeDataForDpi(new HandleRef(null, hWnd), renderer.Class, dpi));
        //        if (hTheme.Handle != IntPtr.Zero)
        //        {
        //            hdc = new HandleRef(null, dc.GetHdc());
        //            DrawThemeBackground(hTheme, hdc, renderer.Part, renderer.State, new COMRECT(bounds), null);
        //        }
        //    }
        //    finally
        //    {
        //        if (hTheme.Handle != IntPtr.Zero)
        //            CloseThemeData(hTheme);

        //        if (hdc.Handle != IntPtr.Zero)
        //            dc.ReleaseHdc();

        //       matrix.Dispose();
        //    }
        //}


        //[DllImport(Uxtheme, CharSet = CharSet.Auto)]
        //private static extern int GetThemeTextExtent(HandleRef hTheme, HandleRef hdc, int iPartId, int iStateId, [MarshalAs(UnmanagedType.LPWStr)] string pszText, int iCharCount, int dwTextFlags,
        //                                             [In] NativeMethods.COMRECT pBoundingRect, [Out] NativeMethods.COMRECT pExtentRect);

        //public static Rectangle GetThemeTextExtent(VisualStyleRenderer renderer, Graphics dc, Rectangle bounds, string textToDraw, TextFormatFlags flags, IntPtr hWnd, int dpi)
        //{
        //    if (dpi == 0 || hWnd == IntPtr.Zero || !OSVersionHelper.IsWindows10CreatorsAdditionOrGreater())
        //    {
        //        return renderer.GetTextExtent(dc, bounds, textToDraw, flags);
        //    }

 
        //    var hdc = new HandleRef(null, dc.GetHdc());
        //    HandleRef hTheme = default(HandleRef);

        //    try
        //    {
        //        NativeMethods.COMRECT pExtentRect = new COMRECT();

        //        var comBounds = bounds.IsEmpty ? null : new NativeMethods.COMRECT(bounds);

        //        hTheme = new HandleRef(renderer, OpenThemeDataForDpi(new HandleRef(null, hWnd), renderer.Class, dpi));
        //        if (hTheme.Handle != IntPtr.Zero)
        //            GetThemeTextExtent(hTheme, hdc, renderer.Part, renderer.State, textToDraw, textToDraw.Length, (int) flags, comBounds, pExtentRect);

        //        return Rectangle.FromLTRB(pExtentRect.Left, pExtentRect.Top, pExtentRect.Right, pExtentRect.Bottom);
        //    }
        //    finally
        //    {
        //        if (hTheme.Handle != IntPtr.Zero)
        //            CloseThemeData(hTheme);
        //        dc.ReleaseHdc();
 
        //    }
        //}


        //[DllImport(Uxtheme, CharSet = CharSet.Auto)]
        //private static extern int DrawThemeText(HandleRef hTheme, HandleRef hdc, int iPartId, int iStateId, [MarshalAs(UnmanagedType.LPWStr)] string pszText, int iCharCount, int dwTextFlags, int dwTextFlags2,
        //                                        [In] NativeMethods.COMRECT pRect);


        //public static Rectangle DrawThemeText(VisualStyleRenderer renderer, Graphics dc, Rectangle bounds, string textToDraw, bool drawDisabled, TextFormatFlags flags, IntPtr hWnd, int dpi,
        //                                      bool returnTextExtents)
        //{
        //    if (dpi == 0 || hWnd == IntPtr.Zero || !OSVersionHelper.IsWindows10CreatorsAdditionOrGreater())
        //    {
        //        renderer.DrawText(dc, bounds, textToDraw, drawDisabled, flags);
        //        return returnTextExtents ? renderer.GetTextExtent(dc, bounds, textToDraw, flags) : Rectangle.Empty;
        //    }

        //    var matrix = dc.Transform;

        //    var hdc = new HandleRef(null, dc.GetHdc());
        //    HandleRef hTheme = default(HandleRef);

        //    try
        //    {
        //        int iDrawDisabled = drawDisabled ? 1 : 0;

        //        bounds.X += (int) (matrix.OffsetX);
        //        bounds.Y += (int) (matrix.OffsetY);

        //        var comBounds = new NativeMethods.COMRECT(bounds);

        //        hTheme = new HandleRef(renderer, OpenThemeDataForDpi(new HandleRef(null, hWnd), renderer.Class, dpi));
        //        if (hTheme.Handle != IntPtr.Zero)
        //            DrawThemeText(hTheme, hdc, renderer.Part, renderer.State, textToDraw, textToDraw.Length, (int) flags, iDrawDisabled, comBounds);

        //        if (!returnTextExtents)
        //            return Rectangle.Empty;

        //        NativeMethods.COMRECT pExtentRect = new COMRECT();
        //        GetThemeTextExtent(hTheme, hdc, renderer.Part, renderer.State, textToDraw, textToDraw.Length, (int) flags, comBounds, pExtentRect);
        //        return Rectangle.FromLTRB(pExtentRect.Left, pExtentRect.Top, pExtentRect.Right, pExtentRect.Bottom);
        //    }
        //    finally
        //    {
        //        if (hTheme.Handle != IntPtr.Zero)
        //            CloseThemeData(hTheme);
        //        dc.ReleaseHdc();

        //        matrix.Dispose();
        //    }
        //}


        public static bool Succeeded(int hr)
        {
            return hr == S_OK;
        }


        [DllImport(Kernel32)]
        public static extern bool ActivateActCtx(IntPtr hActCtx, out IntPtr lpCookie);


        [DllImport(Kernel32)]
        public static extern IntPtr CreateActCtx(ref ACTCTX actctx);


        [DllImport(Kernel32)]
        public static extern bool DeactivateActCtx(int dwFlags, IntPtr lpCookie);


        [DllImport(Kernel32)]
        public static extern bool GetCurrentActCtx(out IntPtr handle);


        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport(User32, CharSet = CharSet.Auto)]
        public static extern bool SystemParametersInfo(SystemParameters uiAction, int uiParam, out int pvParam,
                                                       int fWinIni);


        [DllImport(User32, CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr BeginPaint(HandleRef hWnd, [In, Out] ref PAINTSTRUCT lpPaint);


        [DllImport(User32, CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool EndPaint(HandleRef hWnd, ref PAINTSTRUCT lpPaint);


        [DllImport(User32)]
        public static extern bool GetComboBoxInfo(IntPtr hWnd, ref COMBOBOXINFO pcbi);


        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport("user32.dll")]
        public static extern bool AnimateWindow(IntPtr hwnd, int time, AnimateWindowFlags flags);


        public static int MakeLParam(PointI pt)
        {
            return MakeLParam(Convert.ToInt16(pt.X), Convert.ToInt16(pt.Y));
        }

        public static int MakeLParam(short loWord, short hiWord)
        {
            return (Convert.ToInt32(hiWord) << 0x10) | (Convert.ToInt32(loWord) & 0xffff);
        }

        //Public Shared Function MakeLParamPtr(ByVal loWord As Int16, ByVal hiWord As Int16) As IntPtr
        //	Return New IntPtr(MakeLParam(loWord, hiWord))
        //End Function


        //Public Shared Function MakeWParam(ByVal loWord As Int16, ByVal hiWord As Int16) As IntPtr
        //	Return New IntPtr(MakeLParam(loWord, hiWord))
        //End Function


        public static int LoWord(int word)
        {
            return unchecked((short) word); //returns negitives;

            //return word & 0xffff;
        }

        public static int LoWord(IntPtr word)
        {
            return LoWord(IntPtrToInt32(word));
        }


        public static short HiWord(IntPtr word)
        {
            return unchecked((short) ((uint) word >> 16)); //returns negitives;
            //return HiWord(IntPtrToInt32(word));
        }

        public static short HiWord(int word)
        {
            return Convert.ToInt16(word >> 16);
        }


        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport(User32, CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool GetScrollBarInfo(HandleRef hWnd, int idObject, [In, Out] SCROLLBARINFO psbi);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32, CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool GetScrollInfo(HandleRef hWnd, SBOrientation fnBar, [In, Out] SCROLLINFO si);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32, CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool SetScrollInfo(HandleRef hWnd, SBOrientation fnBar, [In] SCROLLINFO si,
                                                [MarshalAs(UnmanagedType.Bool)] bool fRedraw);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32, CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool HideCaret(HandleRef hWnd);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32, CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool ShowCaret(HandleRef hWnd);


        [DllImport(User32)]
        public static extern int GetGUIThreadInfo(int idThread, ref GUITHREADINFO lpgui);


        [DllImport(Kernel32, EntryPoint = "GetCurrentThreadId", ExactSpelling = true)]
        public static extern int GetCurrentWin32ThreadId();


        [DllImport("kernel32.dll", ExactSpelling = true)]
        public static extern IntPtr GlobalLock(HandleRef handle);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("kernel32.dll", ExactSpelling = true)]
        public static extern bool GlobalUnlock(HandleRef handle);


        //Should return a Size_T (UIntLong) 
        //I dont't expect hGlobals greater than an Int32 (Arrays cannot be resized using an Int64 anyway)
        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return"),
         DllImport("kernel32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern int GlobalSize(HandleRef handle);


        [DllImport(ComCtl32)]
        public static extern int ImageList_DrawEx(IntPtr hIml, int i, IntPtr hdcDst, int x, int y, int dx,
                                                  int dy, int rgbBk, int rgbFg, DrawStyleFlags StyleFlags);


        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern bool BringWindowToTop(IntPtr hWnd);


        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport(User32, CallingConvention = CallingConvention.StdCall)]
        public static extern bool ShowWindow(IntPtr hWnd, ShowWindowFlags nCmdShow);


        [DllImport(Gdi32)]
        public static extern IntPtr CreateCompatibleDC(IntPtr hDc);

        [DllImport(Gdi32)]
        public static extern IntPtr CreateCompatibleBitmap(IntPtr hDc, int cx,int cy);

        
        [DllImport(Gdi32, SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr CreateIC(string lpszDriverName, string lpszDeviceName, string lpszOutput, HandleRef /*DEVMODE*/ lpInitData);


        [DllImport(Gdi32)]
        public static extern int DeleteDC(HandleRef hDc);


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


        //<System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage")> <DllImport(User32, CharSet:=CharSet.Auto, ExactSpelling:=True)> _
        //Public Shared Function GetScrollInfo(ByVal hWnd As HandleRef, ByVal fnBar As Integer, <[In](), Out()> ByVal si As SCROLLINFO) As <MarshalAs(UnmanagedType.Bool)> Boolean
        //End Function


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(Gdi32, CharSet = CharSet.Auto, BestFitMapping = false, ThrowOnUnmappableChar = true)]
        public static extern bool GetTextExtentPoint32(IntPtr hDc, string lpsz, int cbString, ref SizeI lpSize);


        [DllImport(User32)]
        public static extern IntPtr GetWindowDC(HandleRef hWnd);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern int DrawFrameControl(IntPtr hDc, ref RectangleI lpRect, DrawFrameControlType uType,
                                                  DrawFrameControlState uState);


        //<System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage")>
        //<DllImport(User32, SetLastError:=True)>
        //Public Shared Function SetProcessDPIAware() As <MarshalAs(UnmanagedType.Bool)> Boolean
        //End Function

        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport(User32, SetLastError = true)]
        public static extern bool IsProcessDPIAware();


        [DllImport(Kernel32)]
        public static extern int MulDiv(int nNumber, int nNumerator, int nDenominator);


        [DllImport(User32, SetLastError = true)]
        public static extern IntPtr SetFocus(HandleRef hWnd);

        [DllImport(User32)]
        public static extern bool FlashWindowEx(FLASHWINFO pfwi);

        [DllImport(User32)]
        public static extern bool FlashWindow(IntPtr hWnd, bool bInvert);

        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern IntPtr GetWindow(IntPtr hWnd, GetWindowCommand wCmd);


        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern IntPtr GetForegroundWindow();


        [DllImport(User32)]
        public static extern IntPtr GetDCEx(IntPtr hWnd, IntPtr hrgnclip, int fdwOptions);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern IntPtr GetDC(HandleRef hWnd);

        [DllImport(Gdi32)]
        public static extern int GetPixel(IntPtr hdc, int nXPos, int nYPos);
        [DllImport(Gdi32)]
        public static extern int SetPixel(IntPtr hdc, int X, int Y, int crColor);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(Gdi32)]
        public static extern int GetDeviceCaps(IntPtr hDc, int nIndex);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern int ReleaseDC(HandleRef hWnd, IntPtr hDc);


        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage")]
        [DllImport(User32)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy,
                                               SetWindowPosFlags wFlags);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        private static extern bool LockSetForegroundWindow(int uLockCode);

        public static bool LockSetForegroundWindow(bool uLockCode) => LockSetForegroundWindow(uLockCode ? 1 : 2); //1 = Lock, 2 = Unlock


        [DllImport(User32)]
        public static extern bool AttachThreadInput(int idAttach, int idAttachTo, bool fAttach);


        [DllImport(User32)]
        public static extern bool BringWindowToTop(HandleRef hWnd);


        [DllImport(User32)]
        public static extern IntPtr BeginDeferWindowPos(int nNumWindows);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool EndDeferWindowPos(IntPtr hWinPosInfo);


        [DllImport(User32)]
        public static extern IntPtr DeferWindowPos(IntPtr hWinPosInfo, IntPtr hWnd, IntPtr hWndInsertAfter, int x,
                                                   int y, int cx, int cy, SetWindowPosFlags flags);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return"),
         DllImport(User32, EntryPoint = "GetWindowLong", CharSet = CharSet.Auto)]
        private static extern IntPtr GetWindowLongx86(HandleRef hWnd, int nIndex);

        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "2"),
         SuppressMessage("Microsoft.Interoperability", "CA1400:PInvokeEntryPointsShouldExist"),
         DllImport(User32, EntryPoint = "GetWindowLongPtr", CharSet = CharSet.Auto)]
        private static extern IntPtr GetWindowLongX64(HandleRef hWnd, int nIndex);


        public static IntPtr GetWindowLong(HandleRef hWnd, int nIndex)
        {
            if (IntPtr.Size == 4)
            {
                return GetWindowLongx86(hWnd, nIndex);
            }

            return GetWindowLongX64(hWnd, nIndex);
        }

        public static int GetWindowLongInt32(HandleRef hWnd, int nIndex)
        {
            if (IntPtr.Size == 4)
            {
                return IntPtrToInt32(GetWindowLongx86(hWnd, nIndex));
            }

            return IntPtrToInt32(GetWindowLongX64(hWnd, nIndex));
        }


        internal static int IntPtrToInt32(IntPtr value)
        {
            if (4 == IntPtr.Size)
            {
                return value.ToInt32();
            }

            var lval = (long) value;
            lval = Math.Min(int.MaxValue, lval);
            lval = Math.Max(int.MinValue, lval);
            return (int) lval;
        }

        [DllImport(Shlwapi, CharSet = CharSet.Unicode)]
        public static extern int AssocQueryString(uint flags, uint str, string pszAssoc, string pszExtra,
                                                  StringBuilder pszOut, ref int pcchOut);


        [DllImport(Kernel32, CharSet = CharSet.Unicode)]
        public static extern int GetModuleFileName(IntPtr hModule, StringBuilder lpFileName, int nSize);


        [DllImport(Kernel32, CharSet = CharSet.Unicode)]
        public static extern IntPtr GetModuleHandle(string lpModuleName);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32, CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool AdjustWindowRectEx(ref RectangleI lpRect, int dwStyle,
                                                     [MarshalAs(UnmanagedType.Bool)] bool bMenu, int dwExStyle);

        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage")]
        public static RectangleI AdjustWindowRectEx(RectangleI bounds, int dwStyle, bool bMenu, int dwExStyle)
        {
 
            if (AdjustWindowRectEx(ref bounds, dwStyle, bMenu, dwExStyle))
            {
                return bounds;
            }


            return default(RectangleI);
        }

        public static RectangleI AdjustWindowRectEx(HandleRef hWnd, RectangleI bounds, bool bMenu)
        {
        
            var style = GetWindowLongInt32(hWnd, GWL_STYLE);
            var styleEx = GetWindowLongInt32(hWnd, GWL_EXSTYLE);
            if (AdjustWindowRectEx(ref bounds, style, bMenu, styleEx))
            {
                return bounds;
            }
 
            return default(RectangleI);
        }


        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern bool RedrawWindow(IntPtr hWnd, [In] ref RectangleI lprcUpdate, IntPtr hrgnUpdate,
                                               RedrawWindowFlags fuRedraw);

        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern bool RedrawWindow(IntPtr hWnd, IntPtr lprcUpdate, IntPtr hrgnUpdate,
                                               RedrawWindowFlags fuRedraw);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern IntPtr WindowFromPoint(int x, int y);


        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "1"),
         DllImport(User32)]
        public static extern IntPtr ChildWindowFromPoint(IntPtr hWnd, [In] ref PointI pt);

        [DllImport(User32)]
        public static extern IntPtr ChildWindowFromPointEx(IntPtr hWndParent, int ptx, int pty, int uFlags);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern int GetWindowThreadProcessId(HandleRef hWnd, out int lpdwProcessId);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport(User32, EntryPoint = "LoadIcon", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr LoadIcon(HandleRef hInst, IntPtr iconId);


        [DllImport(Kernel32)]
        private static extern int GetCurrentProcessId();


        [SuppressMessage("Microsoft.Usage", "CA1806:DoNotIgnoreMethodResults",
            MessageId = "LevitJames.Libraries.NativeMethods.GetWindowThreadProcessId(System.IntPtr,System.Int32@)")]
        public static bool IsWindowInCurrentProcess(HandleRef hWnd)
        {
            var processID = 0;
            GetWindowThreadProcessId(hWnd, out processID);
            return processID == CurrentProcessId;
        }


        [DllImport(User32)]
        public static extern int ValidateRect(IntPtr hWnd, IntPtr lpRect);

        [DllImport(User32)]
        public static extern int ValidateRect(IntPtr hWnd, ref RectangleI lpRect);

        public static int ValidateRect(IntPtr hWnd, RectangleI rect)
        {
   
            return ValidateRect(hWnd, ref rect);
        }


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool InvalidateRect(IntPtr hWnd, IntPtr lpRect, int bErase);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool InvalidateRect(IntPtr hWnd, [In] ref RectangleI lpRect, int bErase);

        public static bool InvalidateRect(IntPtr hWnd, RectangleI lpRect, int bErase)
        {
   
            return InvalidateRect(hWnd, ref lpRect, bErase);
        }


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool InvalidateRgn(IntPtr hWnd, IntPtr hRgn, int bErase);


        [return: MarshalAs(UnmanagedType.I4)]
        [DllImport(User32)]
        public static extern RegionType GetUpdateRgn(IntPtr hWnd, IntPtr hRgn, int bErase);


        [DllImport(Gdi32)]
        public static extern IntPtr CreateRectRgn(int nLeftRect, int nTopRect, int nRightRect, int nBottomRect);


        [DllImport(Gdi32)]
        public static extern int SetWindowOrgEx(IntPtr hDc, int nX, int nY, out PointI lpPoint);


        [DllImport(Gdi32, EntryPoint = "SetMapMode", CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        public static extern int SetMapMode(HandleRef hDC, int nMapMode);


        [DllImport(Gdi32, CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        public static extern bool SetWindowExtEx(HandleRef hDC, int x, int y, ref SizeI size);

 

        public static bool SetWindowExtEx(HandleRef hDC, int x, int y)
        {
            var sz = new SizeI();
            return SetWindowExtEx(hDC, x, y, ref sz);
        }


        [DllImport(Gdi32, CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        public static extern bool SetViewportExtEx(HandleRef hDC, int x, int y, ref SizeI size);

 
        [DllImport(User32)]
        public static extern IntPtr SetCursor(IntPtr hCursor);

        [DllImport(User32)]
        public static extern IntPtr LoadCursor(int hInstance, int lpCursorName);

        [DllImport(User32, CharSet=CharSet.Auto, ExactSpelling=true)]
        private static extern bool GetCursorPos(ref PointI pt);

        [DllImport(User32)]
        public static extern bool SetCursorPos(int X, int Y);

        public static PointI GetCursorPos()
        {
            var pt = new PointI();
            GetCursorPos(ref pt);
            return pt;
        }
 
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern IntPtr SendMessage(HandleRef hWnd, int uMsg, IntPtr wParam, ref COMBOBOXINFO lParam);

        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern IntPtr SendMessage(HandleRef hWnd, int uMsg, IntPtr wParam, IntPtr lParam);

        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport(User32, CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(HandleRef hWnd, int uMsg, IntPtr wParam, TOOLINFO lParam);

        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern IntPtr SendMessage(HandleRef hWnd, int uMsg, IntPtr wParam, [In] ref PointI lParam);

        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern IntPtr SendMessage(HandleRef hWnd, int uMsg, IntPtr wParam, ref RectangleI lParam);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport(User32, CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(HandleRef hWnd, int uMsg, IntPtr wParam, string lParam);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport(User32, CallingConvention = CallingConvention.StdCall)]
        public static extern IntPtr GetActiveWindow();

        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport(User32, CallingConvention = CallingConvention.StdCall)]
        public static extern bool GetClipCursor(ref RectangleI rect);

        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport("user32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern bool ClipCursor(ref RectangleI rect);

        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport("user32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern bool ClipCursor(IntPtr rect);

        [DllImport(User32, CharSet = CharSet.Auto)]
        public static extern bool PeekMessage([In, Out] ref WindowMessage msg, HandleRef hwnd, int msgMin, int msgMax,
                                              int remove);

        [DllImport(User32)]
        public static extern uint MsgWaitForMultipleObjects(uint nCount, IntPtr[] pHandles, int bWaitAll, uint dwMilliseconds, uint dwWakeMask);

        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return"),
         DllImport(User32)]
        public static extern IntPtr PostMessage(HandleRef hWnd, int uMsg, IntPtr wParam, IntPtr lParam);

        [DllImport(User32, ExactSpelling = true)]
        public static extern IntPtr SetTimer(HandleRef hWnd, IntPtr nIDEvent, uint uElapse, TimerProc lpTimerFunc);

        [DllImport(User32, ExactSpelling = true)]
        public static extern bool KillTimer(HandleRef hWnd, IntPtr nIDEvent);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern IntPtr GetFocus();


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern IntPtr GetDesktopWindow();


        [DllImport(User32, CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        public static extern bool DestroyIcon(IntPtr hIcon);


        [DllImport(User32, CharSet = CharSet.Auto, SetLastError = true, ThrowOnUnmappableChar = true,
            BestFitMapping = false)]
        public static extern IntPtr LoadImage([In] IntPtr hinst, [In] int lpszName, [In] ImageType uType,
                                              [In] int cxDesired, [In] int cyDesired, [In] LoadImageFlags fuLoad);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport(User32, CharSet = CharSet.Auto, BestFitMapping = false, ThrowOnUnmappableChar = true)]
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
        public static extern int GetClassLong(HandleRef hWnd, int nIndex);


        public static Atom GetClassAtom(HandleRef hWnd)
        {
            return new Atom(GetClassLong(hWnd, GCW_ATOM));
        }


        public static Atom GetClassAtomFromName(ref string windowClassName)
        {
            return RegisterClipboardFormatAtom(windowClassName);
        }


        [DllImport(User32, EntryPoint = "RegisterClipboardFormat", CharSet = CharSet.Auto, BestFitMapping = false,
            ThrowOnUnmappableChar = true)]
        private static extern Atom RegisterClipboardFormatAtom(string lpszFormat);

        [DllImport(User32, CharSet = CharSet.Auto, BestFitMapping = false, ThrowOnUnmappableChar = true)]
        public static extern int RegisterClipboardFormat(string lpszFormat);


        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        public static extern int GetWindowTextLength(HandleRef hWnd);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport(User32, CharSet = CharSet.Auto, BestFitMapping = false, ThrowOnUnmappableChar = true)]
        public static extern int GetWindowText(HandleRef hWnd,
                                               [MarshalAs(UnmanagedType.LPTStr)] StringBuilder lpString, int cch);

        public static string GetWindowText(IntPtr hWnd)
        {
            var capacity = GetWindowTextLength(new HandleRef(wrapper: null, handle: hWnd)) * 2;
            if (capacity > 0)
            {
                var text = new StringBuilder(capacity);
                GetWindowText(new HandleRef(wrapper: null, handle: hWnd), text, text.Capacity);
                return text.ToString();
            }

            return null;
        }


        [DllImport(User32, CharSet = CharSet.Auto, BestFitMapping = false, ThrowOnUnmappableChar = true)]
        public static extern IntPtr FindWindowEx(IntPtr hWndParent, IntPtr hWndChildAfter, string lpszClass,
                                                 string lpszWindow);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern int SetParent(IntPtr hWndChild, IntPtr hWndNewParent);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport(User32, CharSet = CharSet.Auto, BestFitMapping = false, ThrowOnUnmappableChar = true,
             SetLastError = true)]
        public static extern int GetProp(HandleRef hWnd, string lpString);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport(User32, CharSet = CharSet.Auto, BestFitMapping = false, ThrowOnUnmappableChar = true,
             SetLastError = true)]
        public static extern int SetProp(HandleRef hWnd, string lpString, int value);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport(User32, CharSet = CharSet.Auto, BestFitMapping = false, ThrowOnUnmappableChar = true,
             SetLastError = true)]
        public static extern int RemoveProp(HandleRef hWnd, string lpString);


        [DllImport(User32, SetLastError = true)]
        public static extern bool PrintWindow(HandleRef hwnd, IntPtr hDC, uint nFlags);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage")]
        [DllImport(User32)]
        public static extern IntPtr GetParent(IntPtr hWnd);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage")]
        [DllImport(User32)]
        public static extern IntPtr GetTopWindow(IntPtr hWnd);


        [DllImport(User32)]
        public static extern int UpdateWindow(IntPtr hWnd);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport(Gdi32, ExactSpelling = true)]
        public static extern int GetDCOrgEx(IntPtr hWnd, ref PointI lpPoint);


        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport(User32, ExactSpelling = true)]
        public static extern bool UpdateLayeredWindow(HandleRef hwnd, IntPtr hdcDst,
                                                      [MarshalAs(UnmanagedType.AsAny)] object pptDst,
                                                      [MarshalAs(UnmanagedType.AsAny)] object psize, IntPtr hdcSrc,
                                                      [MarshalAs(UnmanagedType.AsAny)] object pprSrc, int crKey,
                                                      [MarshalAs(UnmanagedType.AsAny), In] object pblend, int dwFlags);


        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "2"),
         SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport(User32, CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        public static extern bool SetLayeredWindowAttributes(HandleRef hwnd, int crKey, byte bAlpha, int dwFlags);


        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        private static extern bool GetWindowRect(HandleRef hWnd, ref RectangleI lpRect);

        //Public Shared Function GetWindowRect(ByVal hWnd As IntPtr) As Rectangle
        //	Dim rc As RectangleI
        //	If GetWindowRect(hWnd, rc) Then
        //		Return New Rectangle(rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top)
        //	End If
        //	Return Nothing
        //End Function
        public static RectangleI GetWindowRect(HandleRef hWnd)
        {
            var rc = new RectangleI();
            if (GetWindowRect(hWnd, ref rc))
            {
                return rc;
            }

            return default(RectangleI);
        }


        //Note this only does one rectangle at the moment
        [DllImport(User32)]
        private static extern int MapWindowPoints(IntPtr hWndFrom, IntPtr hWndTo, ref RectangleI lpRect, uint cPoints);

        [DllImport(User32)]
        private static extern int MapWindowPoints(IntPtr hWndFrom, IntPtr hWndTo, ref PointI lpPoint, uint cPoints);

        public static RectangleI MapWindowPoints(IntPtr hWndFrom, IntPtr hWndTo, RectangleI rc)
        {
            _ = MapWindowPoints(hWndFrom, hWndTo, ref rc, cPoints: 2); //2 as there are two points in a rect
            return rc;
        }

        public static PointI MapWindowPoints(IntPtr hWndFrom, IntPtr hWndTo, PointI pt)
        {
            _ = MapWindowPoints(hWndFrom, hWndTo, ref pt, cPoints: 1);
            return pt;
        }


        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        private static extern bool GetClientRect(HandleRef hWnd, ref RectangleI lpRect);

        public static RectangleI GetClientRect(HandleRef hWnd)
        {
            var rc = new RectangleI();
            if (GetClientRect(hWnd, ref rc))
            {
                return rc;
            }

            return default(RectangleI);
        }


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return"),
         DllImport(User32, EntryPoint = "SetWindowLong", CharSet = CharSet.Auto)]
        private static extern IntPtr SetWindowLongx86(HandleRef hWnd, int nIndex, int dwNewLong);

        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "2"),
         SuppressMessage("Microsoft.Interoperability", "CA1400:PInvokeEntryPointsShouldExist"),
         DllImport(User32, EntryPoint = "SetWindowLongPtr", CharSet = CharSet.Auto)]
        private static extern IntPtr SetWindowLongX64(HandleRef hWnd, int nIndex, IntPtr dwNewLong);

        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return"),
         DllImport(User32, EntryPoint = "SetWindowLong", CharSet = CharSet.Auto)]
        private static extern IntPtr SetWindowLongx86(HandleRef hWnd, int nIndex, HandleRef dwNewLong);

        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "2"),
         SuppressMessage("Microsoft.Interoperability", "CA1400:PInvokeEntryPointsShouldExist"),
         DllImport(User32, EntryPoint = "SetWindowLongPtr", CharSet = CharSet.Auto)]
        private static extern IntPtr SetWindowLongX64(HandleRef hWnd, int nIndex, HandleRef dwNewLong);


        public static IntPtr SetWindowLong(HandleRef hWnd, int nIndex, int dwNewLong)
        {
            if (IntPtr.Size == 4)
            {
                return SetWindowLongx86(hWnd, nIndex, dwNewLong);
            }

            return SetWindowLongX64(hWnd, nIndex, new IntPtr(dwNewLong));
        }

        public static IntPtr SetWindowLong(HandleRef hWnd, int nIndex, IntPtr dwNewLong)
        {
            if (IntPtr.Size == 4)
            {
                return SetWindowLongx86(hWnd, nIndex, dwNewLong.ToInt32());
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


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(Kernel32)]
        public static extern int GetCurrentThreadId();


        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern bool IsWindow(HandleRef hWnd);


        [DllImport(User32)]
        public static extern IntPtr SetCapture(IntPtr hWnd);

        [DllImport(User32)]
        public static extern bool ReleaseCapture();


        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern bool ShowOwnedPopups(HandleRef hWnd, [MarshalAs(UnmanagedType.Bool)] bool fShow);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool IsIconic(IntPtr hWnd);


        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern bool IsWindowVisible(HandleRef hWnd);


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


        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern bool IsWindowEnabled(HandleRef hWnd);


        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern bool EnableWindow(IntPtr hWnd, [MarshalAs(UnmanagedType.Bool)] bool fEnable);

        [return: MarshalAs(UnmanagedType.Bool)]
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern bool EnumChildWindows(HandleRef hWndParent, EnumChildProc lpEnumFunc, IntPtr lParam);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(User32)]
        public static extern bool EnumThreadWindows(int dwThreadID, EnumThreadWndProc lpfn, HandleRef lParam);

 

        [DllImport(User32)]
        public static extern IntPtr GetAncestor(IntPtr hWnd, GAFlags gaFlags);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), DllImport(User32)]
        public static extern int SetActiveWindow(HandleRef hWnd);


        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"),
         DllImport(User32, CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr GetCapture();

        [DllImport(MsImg32)]
        public static extern int AlphaBlend(IntPtr hdcDest, int nXOriginDest, int nYOriginDest, int nWidthDest,
                                            int nHeightDest, IntPtr hdcSrc, int nXOriginSrc, int nYOriginSrc,
                                            int nWidthSrc, int nHeightSrc, BLENDFUNCTION lBlendFunction);

        [DllImport(User32, CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool CloseWindow(IntPtr handle);

        [DllImport(OleAut32, SetLastError = false, ExactSpelling = true)]
        public static extern int OleLoadPicture(IStream lpstream, int lSize, [MarshalAs(UnmanagedType.Bool)] bool fRunmode, in Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object lplpvObj);

        public static void ThrowWin32Error()
        {
            ThrowWin32Error(message: null);
        }

        [SecurityCritical, SuppressMessage("Microsoft.Usage", "CA2201:DoNotRaiseReservedExceptionTypes")]
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
        public static extern int AccessibleObjectFromWindow(IntPtr hWnd, int dwid, [In] ref Guid riid,
                                                            [MarshalAs(UnmanagedType.IDispatch)] ref object ppvobject);

         

        [DllImport(Uxtheme, CharSet = CharSet.Unicode)]
        public static extern int SetWindowTheme(IntPtr hWnd, string pszSubAppName, string pszSubIdList);


        [DllImport(User32, EntryPoint = "SetClassLong")]
        public static extern uint SetClassLongPtr32(IntPtr hWnd, int nIndex, uint dwNewLong);

        [DllImport(User32, EntryPoint = "SetClassLongPtr")]
        private static extern IntPtr SetClassLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);


        public static IntPtr SetClassLong(IntPtr hWnd, int nIndex, IntPtr dwNewLong)
        {
            //check for x64
            if (IntPtr.Size == 8)
                return SetClassLongPtr64(hWnd, nIndex, dwNewLong);
            else
            {
                var uintvalue = unchecked((uint)dwNewLong.ToInt32());
                return new IntPtr(SetClassLongPtr32(hWnd, nIndex, uintvalue));
            }
                
        }


        [DllImport(GdiPlus)]
        public static extern int GdipWindingModeOutline(HandleRef path, IntPtr matrix, float flatness);


        [DllImport(User32)]
        public static extern IntPtr GetDlgItem(HandleRef hwnd, int id);


        [DllImport(Shell32)]
        public static extern int SHGetStockIconInfo(int siid, SHGSI uFlags, ref SHSTOCKICONINFO psii);

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        public static extern int SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbFileInfo, NativeMethods.SHGFI uFlags);

        [DllImport(Kernel32, BestFitMapping = false, ExactSpelling = true, CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr LoadLibraryExW([In] string lpLibFileName, [In] IntPtr hFile, [In] uint dwFlags);

   
        [DllImport(Kernel32, CharSet = CharSet.Ansi, ExactSpelling = true)]
        public static extern IntPtr GetProcAddress(HandleRef hModule, string lpProcName);

        public enum SHIL { Large = 0x0, Small = 0x1, ExtraLarge = 0x2, Jumbo = 0x4}
        [DllImport("shell32.dll", EntryPoint = "#727")]
        public extern static int SHGetImageList(SHIL iImageList, ref Guid riid, ref IImageList ppv);

        [DllImport(Ole32)]
        public static extern int OleSetClipboard(IntPtr pDataObj);


        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport(Kernel32, CharSet = CharSet.Auto)]
        internal static extern bool FreeLibrary(IntPtr hModule);


        /// <summary>
        ///     A method for returning the Version of a dll module.
        /// </summary>
        /// <param name="moduleName"></param>
        /// <returns>The Version information for the module, or null if no information is found.</returns>
        /// <remarks>
        ///     If the module is not already loaded it will be loaded and then freed. The method will first try and get the version
        ///     using the DllGetVersion export. If the DllGetVersion export is not implemented then the file version information is
        ///     read.
        /// </remarks>
        public static Version GetModuleVersion(string moduleName)
        {
            var freeLib = false;
            Version version = null;
            var hModule = GetModuleHandle(moduleName);

            if (hModule == IntPtr.Zero)
            {
                hModule = LoadLibraryExW(moduleName, IntPtr.Zero, dwFlags: 0);
                freeLib = true;
            }

            if (hModule != IntPtr.Zero)
            {
                var pProc = GetProcAddress(new HandleRef(wrapper: null, handle: hModule), "DllGetVersion");
                if (pProc != IntPtr.Zero)
                {
                    var pFunc =
                        (DllGetVersionProc) Marshal.GetDelegateForFunctionPointer(pProc, typeof(DllGetVersionProc));
                    var vi = new DllVersionInfo();
                    if (pFunc.Invoke(vi) == 0)
                    {
                        version = new Version(vi.dwMajorVersion, vi.dwMinorVersion, vi.dwBuildNumber);
                    }
                }
                else
                {
                    var buffer = new StringBuilder(capacity: 260);
                    var r = GetModuleFileName(hModule, buffer, buffer.Capacity);
                    if (r > 0 && File.Exists(buffer.ToString()))
                    {
                        var fv = FileVersionInfo.GetVersionInfo(buffer.ToString());
                        version = new Version(fv.FileMajorPart, fv.FileMinorPart, fv.FileBuildPart);
                    }
                }
            }

            if (freeLib && hModule != IntPtr.Zero)
            {
                FreeLibrary(hModule);
            }

            return version;
        }


        [DllImport(ComCtl32)]
        internal static extern int SetWindowSubclass(IntPtr hWnd, SubClassProcDelegate newProc, UIntPtr uIdSubclass,
                                                     IntPtr dwRefData);


        [DllImport(ComCtl32)]
        internal static extern int RemoveWindowSubclass(HandleRef hWnd, SubClassProcDelegate newProc, UIntPtr uIdSubclass);


        [DllImport(ComCtl32)]
        internal static extern IntPtr DefSubclassProc(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        [DllImport(User32, CharSet = CharSet.Auto)]
        public static extern IntPtr CallWindowProc(IntPtr lpPrevWndFunc, IntPtr hwnd, Int32 msg, IntPtr wParam, IntPtr lParam);


        [DllImport(Ole32)]
        public static extern int CoRegisterMessageFilter(HandleRef lpMessageFilter, out IntPtr lplpMessageFilter);


        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return"),
         DllImport(User32, CharSet = CharSet.Auto)]
        public static extern IntPtr CallNextHookEx(WindowsHookSafeHandle hook, int nCode, IntPtr wParam, IntPtr lParam);


        [DllImport(User32)]
        public static extern bool TranslateMessage([In, Out] ref WindowMessage msg);

        [DllImport(User32)]
        public static extern int DispatchMessage([In] ref WindowMessage msg);

        [Flags]
        public enum ChooseColorFlags
        {
	        RgbInt = 0x1,
	        FullOpen = 0x2,
	        PreventFullOpen = 0x4,
	        ShowHelp = 0x8,
	        EnableHook = 0x10,
	        EnableTemplate = 0x20,
	        EnableTemplateHandle = 0x40,
	        SolidColor = 0x080,
	        AnyColor = 0x100,

        }


        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public class CHOOSECOLOR
        {
	        public int lStructSize = Marshal.SizeOf(typeof(NativeMethods.CHOOSECOLOR));
	        public IntPtr hwndOwner = IntPtr.Zero;
	        public IntPtr hInstance = IntPtr.Zero;
	        public int rgbResult;
	        public IntPtr lpCustColors = IntPtr.Zero;
	        public ChooseColorFlags Flags;
	        public IntPtr lCustData = IntPtr.Zero;
	        public IntPtr lpfnHook = IntPtr.Zero;
	        public string lpTemplateName;
        }


        [DllImport(Comdlg32, CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool ChooseColor([In, Out] CHOOSECOLOR lpChooseColor);

        [StructLayout(LayoutKind.Sequential)]
        public struct ACTCTX
        {
            public int cbSize;
            public uint dwFlags;
            public string lpSource;
            public ushort wProcessorArchitecture;
            public ushort wLangId;
            public string lpAssemblyDirectory;
            public IntPtr lpResourceName;
            public string lpApplicationName;
            public IntPtr hModule;
        }


        [StructLayout(LayoutKind.Sequential)]
        public struct INITCOMMONCONTROLSEX
        {
            [DllImport("comctl32")]
            public static extern int InitCommonControlsEx(ref INITCOMMONCONTROLSEX lpInitcommoncontrolsex);

            public int dwSize; //size of this structure
            public int dwICC; //flags indicating which classes to be initialized

            public INITCOMMONCONTROLSEX(int flags) : this()
            {
                dwSize = Marshal.SizeOf(typeof(INITCOMMONCONTROLSEX));
                dwICC = flags;
            }

            public static void InitCommonControls(int flags)
            {
                var t = new INITCOMMONCONTROLSEX(flags);
                try
                {
                    InitCommonControlsEx(ref t);
                }
                catch { }
            }
        }


        [StructLayout(LayoutKind.Sequential)]
        public struct NCCALCSIZE_PARAMS
        {
            public RectangleI rgrc0;
            public RectangleI rgrc1;
            public RectangleI rgrc2;
            public IntPtr lppos;
        }


        [StructLayout(LayoutKind.Sequential)]
        public class MINMAXINFO
        {
            public PointI ptReserved;
            public PointI ptMaxSize;
            public PointI ptMaxPosition;
            public PointI ptMinTrackSize;
            public PointI ptMaxTrackSize;
        }


        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public class LOGFONT
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




        [StructLayout(LayoutKind.Explicit, Size = 4)]
        public struct Atom : IComparable, IFormattable, IConvertible, IComparable<int>, IEquatable<int>
        {
            public const int MaxValue = 2147483647;
            public const int MinValue = -2147483648;

            [FieldOffset(offset: 0)] public int _value;

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

            [SuppressMessage("Microsoft.Globalization", "CA1305:SpecifyIFormatProvider",
                MessageId = "System.Int32.ToString(System.String)")]
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

            [SuppressMessage("Microsoft.Globalization", "CA1305:SpecifyIFormatProvider",
                MessageId = "System.Int32.Parse(System.String)")]
            public static Atom Parse(string s)
            {
                return new Atom(int.Parse(s));
            }

            [SuppressMessage("Microsoft.Globalization", "CA1305:SpecifyIFormatProvider",
                MessageId = "System.Int32.Parse(System.String,System.Globalization.NumberStyles)")]
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
                var value = 0;
                result = default(Atom);
                var r = int.TryParse(s, out value);
                if (r)
                {
                    result = new Atom(value);
                }

                return r;
            }

            public static bool TryParse(string s, NumberStyles style, IFormatProvider provider, out Atom result)
            {
                var value = 0;
                result = default(Atom);
                var r = int.TryParse(s, style, provider, out value);
                if (r)
                {
                    result = new Atom(value);
                }

                return r;
            }

            public TypeCode GetTypeCode()
            {
                return TypeCode.Int32;
            }

            bool IConvertible.ToBoolean(IFormatProvider provider)
            {
                return ToBoolean(provider);
            }

            private bool ToBoolean(IFormatProvider provider)
            {
                return Convert.ToBoolean(_value);
            }

            char IConvertible.ToChar(IFormatProvider provider)
            {
                return ToChar(provider);
            }

            private char ToChar(IFormatProvider provider)
            {
                return Convert.ToChar(_value);
            }

            sbyte IConvertible.ToSByte(IFormatProvider provider)
            {
                return ToSByte(provider);
            }

            private sbyte ToSByte(IFormatProvider provider)
            {
                return Convert.ToSByte(_value);
            }

            byte IConvertible.ToByte(IFormatProvider provider)
            {
                return ToByte(provider);
            }

            private byte ToByte(IFormatProvider provider)
            {
                return Convert.ToByte(_value);
            }

            short IConvertible.ToInt16(IFormatProvider provider)
            {
                return ToInt16(provider);
            }

            private short ToInt16(IFormatProvider provider)
            {
                return Convert.ToInt16(_value);
            }

            ushort IConvertible.ToUInt16(IFormatProvider provider)
            {
                return ToUInt16(provider);
            }

            private ushort ToUInt16(IFormatProvider provider)
            {
                return Convert.ToUInt16(_value);
            }

            int IConvertible.ToInt32(IFormatProvider provider)
            {
                return ToInt32(provider);
            }

            private int ToInt32(IFormatProvider provider)
            {
                return _value;
            }

            uint IConvertible.ToUInt32(IFormatProvider provider)
            {
                return ToUInt32(provider);
            }

            private uint ToUInt32(IFormatProvider provider)
            {
                return Convert.ToUInt32(_value);
            }

            long IConvertible.ToInt64(IFormatProvider provider)
            {
                return ToInt64(provider);
            }

            private long ToInt64(IFormatProvider provider)
            {
                return Convert.ToInt64(_value);
            }

            ulong IConvertible.ToUInt64(IFormatProvider provider)
            {
                return ToUInt64(provider);
            }

            private ulong ToUInt64(IFormatProvider provider)
            {
                return Convert.ToUInt64(_value);
            }

            float IConvertible.ToSingle(IFormatProvider provider)
            {
                return ToSingle(provider);
            }

            private float ToSingle(IFormatProvider provider)
            {
                return Convert.ToSingle(_value);
            }

            double IConvertible.ToDouble(IFormatProvider provider)
            {
                return ToDouble(provider);
            }

            private double ToDouble(IFormatProvider provider)
            {
                return Convert.ToDouble(_value);
            }

            decimal IConvertible.ToDecimal(IFormatProvider provider)
            {
                return ToDecimal(provider);
            }

            private decimal ToDecimal(IFormatProvider provider)
            {
                return Convert.ToDecimal(_value);
            }

            DateTime IConvertible.ToDateTime(IFormatProvider provider)
            {
                return ToDateTime(provider);
            }

            private DateTime ToDateTime(IFormatProvider provider)
            {
                return ((IConvertible) _value).ToDateTime(provider);
            }

            object IConvertible.ToType(Type type, IFormatProvider provider)
            {
                return ToType(type, provider);
            }

            private object ToType(Type type, IFormatProvider provider)
            {
                return ((IConvertible) _value).ToType(type, provider);
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

            public static bool operator !=(Atom value1, int value2) => value1._value != value2;

            public static bool operator !=(int value1, Atom value2)
            {
                return value1 != value2._value;
            }
        }


        [StructLayout(LayoutKind.Sequential)]
        public struct PAINTSTRUCT
        {
            public IntPtr hdc;
            public bool fErase;
            public int rcPaint_left;
            public int rcPaint_top;
            public int rcPaint_right;
            public int rcPaint_bottom;
            public bool fRestore;
            public bool fIncUpdate;
            public int reserved1;
            public int reserved2;
            public int reserved3;
            public int reserved4;
            public int reserved5;
            public int reserved6;
            public int reserved7;
            public int reserved8;
        }


        [StructLayout(LayoutKind.Sequential)]
        public class COMSIZE
        {
            public int Width;
            public int Height;
      
            public SizeI ToSize => new SizeI(Width, Height);
        }

        [StructLayout(LayoutKind.Sequential)]
        public class COMRECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
            public COMRECT() { }

            public COMRECT(RectangleI rc)
            {
                Left = rc.Left;
                Top = rc.Top;
                Right = rc.Right;
                Bottom = rc.Bottom;
            }
        }


        [SuppressMessage("Microsoft.Design", "CA1049:TypesThatOwnNativeResourcesShouldBeDisposable"),
         StructLayout(LayoutKind.Sequential)]
        public struct NMHDR
        {
            [SuppressMessage("Microsoft.Reliability", "CA2006:UseSafeHandleToEncapsulateNativeResources")]
            public IntPtr hwndFrom;

            [SuppressMessage("Microsoft.Reliability", "CA2006:UseSafeHandleToEncapsulateNativeResources")]
            public IntPtr idFrom;

            public int code;
        }


        [StructLayout(LayoutKind.Sequential)]
        public struct COMBOBOXINFO
        {
            public int cbSize;
            public RectangleI rcItem;
            public RectangleI rcButton;
            public ComboBoxButtonState buttonState;
            public IntPtr hwndCombo;
            public IntPtr hwndEdit;
            public IntPtr hwndList;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        internal class TOOLINFO
        {
            public int cbSize;
            public int uFlags;
            public IntPtr hWnd;
            public IntPtr uID;
            public RectangleI winRect;
            public IntPtr hInst;
            public IntPtr lpszText;
            public IntPtr lParam;
            public IntPtr lpReserved;

            public static int GetToolTipInfoSize()
            {
                var v = GetModuleVersion(ComCtl32);
                if (v != null && v.Major >= 6)
                {
                    return Marshal.SizeOf(typeof(TOOLINFO));
                }
                return TTTOOLINFO_V2_SIZE;
            }
        }


        [StructLayout(LayoutKind.Sequential)]
        public class SCROLLINFO
        {
            public int cbSize = Marshal.SizeOf(typeof(SCROLLINFO));
            public ScrollInfoMask fMask;
            public int nMin;
            public int nMax;
            public int nPage;
            public int nPos;
            public int nTrackPos;
        }


        [StructLayout(LayoutKind.Sequential)]
        public class SCROLLBARINFO
        {
            public int cbSize = Marshal.SizeOf(typeof(SCROLLBARINFO));
            public RectangleI rcScrollBar;
            public int dxyLineButton;
            public int xyThumbTop;
            public int xyThumbBottom;
            public int reserved;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 6)]
            public int[] rgstate;
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

            public PointI Location
            {
                get { return new PointI(x, y); }
                set
                {
                    x = value.X;
                    y = value.Y;
                }
            }
             
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

            public BLENDFUNCTION(byte blendOp, byte blendFlags, byte sourceConstantAlpha, byte alphaFormat) : this()
            {
                BlendOp = blendOp;
                BlendFlags = blendFlags;
                SourceConstantAlpha = sourceConstantAlpha;
                AlphaFormat = alphaFormat;
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
            public RectangleI rcCaret;

            public static GUITHREADINFO NewGUITHREADINFO()
            {
                var guiTHREADINFO = new GUITHREADINFO();
                guiTHREADINFO.cbSize = Marshal.SizeOf(typeof(GUITHREADINFO));
                return guiTHREADINFO;
            }
        }


        [StructLayout(LayoutKind.Sequential)]
        public class FLASHWINFO
        {
            public int cbSize = Marshal.SizeOf(typeof(FLASHWINFO));
            public int dwFlags;
            public int dwTimeout;
            public IntPtr hwnd;
            public int uCount;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct SHFILEINFO {
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
 
        }
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct SHSTOCKICONINFO
        {
            public int cbSize;
            public IntPtr hIcon;
            public int iSysIconIndex;
            public int iIcon;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = MAX_PATH)]
            public string szPath;
        }
     

        [DllImport(User32)]
        public static extern bool SetWindowPlacement(IntPtr hWnd, [In] WindowPlacement lpwndpl);

        [DllImport(User32)]
        public static extern bool GetWindowPlacement(IntPtr hWnd, [In, Out] WindowPlacement lpwndpl);


        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private class DllVersionInfo
        {
            public int cbSize = Marshal.SizeOf(typeof(DllVersionInfo));
            public int dwMajorVersion;
            public int dwMinorVersion;
            public int dwBuildNumber;
            public int dwPlatformID;
        }


        private delegate int DllGetVersionProc(DllVersionInfo v);


        internal delegate IntPtr SubClassProcDelegate(
            IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam, UIntPtr uIdSubclass, IntPtr dwRefData);


        //<DllImport(ole32, ExactSpelling:=True)> _
        //Public Shared Function CoRegisterMessageFilter(ByVal newFilter As HandleRef, ByRef oldMsgFilter As IntPtr) As Integer
        //End Function

        //Public Shared Function Succeeded(ByVal hr As Integer) As Boolean
        //	Return (hr >= 0)
        //End Function
        [ComImport, Guid("6D5140C1-7436-11CE-8034-00AA006009FA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IOleServiceProvider
        {
            [PreserveSig]
            int QueryService([In] ref Guid guidService, [In] ref Guid riid, out IntPtr ppvObject);
        }


        public class WindowsHookSafeHandle : SafeHandle
        {
            // ReSharper disable once NotAccessedField.Local
            private fnHookProc _fnHookProc;
            //Must keep a ref to this so we don't get a CallbackOnCollectedDelegate Exception


            public WindowsHookSafeHandle() : base(IntPtr.Zero, ownsHandle: true) { }


            public override bool IsInvalid
            {
                [DebuggerStepThrough] get { return handle == IntPtr.Zero; }
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

            public enum WmSizingParameters
            {
                Left =1,
                Right =2,
                Top=3,
                TopLeft =4,
                TopRight=5,
                Bottom = 6,
                BottomLeft = 7,
                BottomRight = 8
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


    }
}