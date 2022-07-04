// © Copyright 2018 Levit & James, Inc.
#define DEBUG_WIN_MSG

using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
 
// ReSharper disable InconsistentNaming

namespace LevitJames.Win32
{
    [CompilerGenerated]
    public enum NotificationMessages
    {
        NM_FIRST = 0 - 0,
        NM_CUSTOMDRAW = NM_FIRST - 12,
        NM_NCHITTEST = NM_FIRST - 14
    }


    [CompilerGenerated]
    public enum ReflectedMessages
    {
        WM_REFLECTED = WindowMessages.WM_USER + 0x1C00,

        WM_COMMAND = WM_REFLECTED + WindowMessages.WM_COMMAND,
        WM_CTLCOLORBTN = WM_REFLECTED + WindowMessages.WM_CTLCOLORBTN,
        WM_CTLCOLOREDIT = WM_REFLECTED + WindowMessages.WM_CTLCOLOREDIT,
        WM_CTLCOLORDLG = WM_REFLECTED + WindowMessages.WM_CTLCOLORDLG,
        WM_CTLCOLORLISTBOX = WM_REFLECTED + WindowMessages.WM_CTLCOLORLISTBOX,
        WM_CTLCOLORMSGBOX = WM_REFLECTED + WindowMessages.WM_CTLCOLORMSGBOX,
        WM_CTLCOLORSCROLLBAR = WM_REFLECTED + WindowMessages.WM_CTLCOLORSCROLLBAR,
        WM_CTLCOLORSTATIC = WM_REFLECTED + WindowMessages.WM_CTLCOLORSTATIC,
        WM_CTLCOLOR = WM_REFLECTED + WindowMessages.WM_CTLCOLOR,
        WM_DRAWITEM = WM_REFLECTED + WindowMessages.WM_DRAWITEM,
        WM_MEASUREITEM = WM_REFLECTED + WindowMessages.WM_MEASUREITEM,
        WM_DELETEITEM = WM_REFLECTED + WindowMessages.WM_DELETEITEM,
        WM_VKEYTOITEM = WM_REFLECTED + WindowMessages.WM_VKEYTOITEM,
        WM_CHARTOITEM = WM_REFLECTED + WindowMessages.WM_CHARTOITEM,
        WM_COMPAREITEM = WM_REFLECTED + WindowMessages.WM_COMPAREITEM,
        WM_HSCROLL = WM_REFLECTED + WindowMessages.WM_HSCROLL,
        WM_VSCROLL = WM_REFLECTED + WindowMessages.WM_VSCROLL,
        WM_PARENTNOTIFY = WM_REFLECTED + WindowMessages.WM_PARENTNOTIFY,
        WM_NOTIFY = WM_REFLECTED + WindowMessages.WM_NOTIFY
    }


    [CompilerGenerated]
    public static class WindowMessages
    {
        // ReSharper disable InconsistentNaming
        public const int WM_ACTIVATE = 0x6;
        public const int WM_ACTIVATEAPP = 0x1C;
        public const int WM_AFXFIRST = 0x360;
        public const int WM_AFXLAST = 0x37F;
        public const int WM_APP = 0x8000;
        public const int WM_ASKCBFORMATNAME = 0x30C;
        public const int WM_CANCELJOURNAL = 0x4B;
        public const int WM_CANCELMODE = 0x1F;
        public const int WM_CAPTURECHANGED = 0x215;
        public const int WM_CHANGECBCHAIN = 0x30D;
        public const int WM_CHAR = 0x102;
        public const int WM_CHARTOITEM = 0x2F;
        public const int WM_CHILDACTIVATE = 0x22;
        public const int WM_CHOOSEFONT_GETLOGFONT = WM_USER + 1;
        public const int WM_CHOOSEFONT_SETFLAGS = WM_USER + 102;
        public const int WM_CHOOSEFONT_SETLOGFONT = WM_USER + 101;
        public const int WM_CLEAR = 0x303;
        public const int WM_CLOSE = 0x10;
        public const int WM_COMMAND = 0x111;
        public const int WM_COMMNOTIFY = 0x44;
        public const int WM_COMPACTING = 0x41;
        public const int WM_COMPAREITEM = 0x39;
        public const int WM_CONTEXTMENU = 0x7B;
        public const int WM_CONVERTREQUESTEX = 0x108;
        public const int WM_COPY = 0x301;
        public const int WM_COPYDATA = 0x4A;
        public const int WM_CREATE = 0x1;
        public const int WM_CTLCOLOR = 0x19;
        public const int WM_CTLCOLORBTN = 0x135;
        public const int WM_CTLCOLORDLG = 0x136;
        public const int WM_CTLCOLOREDIT = 0x133;
        public const int WM_CTLCOLORLISTBOX = 0x134;
        public const int WM_CTLCOLORMSGBOX = 0x132;
        public const int WM_CTLCOLORSCROLLBAR = 0x137;
        public const int WM_CTLCOLORSTATIC = 0x138;
        public const int WM_CUT = 0x300;
        public const int WM_DDE_ACK = WM_DDE_FIRST + 4;
        public const int WM_DDE_ADVISE = WM_DDE_FIRST + 2;
        public const int WM_DDE_DATA = WM_DDE_FIRST + 5;
        public const int WM_DDE_EXECUTE = WM_DDE_FIRST + 8;
        public const int WM_DDE_FIRST = 0x3E0;
        public const int WM_DDE_INITIATE = WM_DDE_FIRST;
        public const int WM_DDE_LAST = WM_DDE_FIRST + 8;
        public const int WM_DDE_POKE = WM_DDE_FIRST + 7;
        public const int WM_DDE_REQUEST = WM_DDE_FIRST + 6;
        public const int WM_DDE_TERMINATE = WM_DDE_FIRST + 1;
        public const int WM_DDE_UNADVISE = WM_DDE_FIRST + 3;
        public const int WM_DEADCHAR = 0x103;
        public const int WM_DELETEITEM = 0x2D;
        public const int WM_DESTROY = 0x2;
        public const int WM_DESTROYCLIPBOARD = 0x307;
        public const int WM_DEVICECHANGE = 0x219;
        public const int WM_DEVMODECHANGE = 0x1B;
        public const int WM_DISPLAYCHANGE = 0x7E;
        public const int WM_DRAWCLIPBOARD = 0x308;
        public const int WM_DRAWITEM = 0x2B;
        public const int WM_DROPFILES = 0x233;
        public const int WM_ENABLE = 0xA;
        public const int WM_ENDSESSION = 0x16;
        public const int WM_ENTERIDLE = 0x121;
        public const int WM_ENTERMENULOOP = 0x211;
        public const int WM_ERASEBKGND = 0x14;
        public const int WM_EXITMENULOOP = 0x212;
        public const int WM_FONTCHANGE = 0x1D;
        public const int WM_GETDLGCODE = 0x87;
        public const int WM_GETFONT = 0x31;
        public const int WM_GETHOTKEY = 0x33;
        public const int WM_GETICON = 0x7F;
        public const int WM_GETMINMAXINFO = 0x24;
        public const int WM_GETTEXT = 0xD;
        public const int WM_GETTEXTLENGTH = 0xE;
        public const int WM_HANDHELDFIRST = 0x358;
        public const int WM_HANDHELDLAST = 0x35F;
        public const int WM_HELP = 0x53;
        public const int WM_HOTKEY = 0x312;
        public const int WM_HSCROLL = 0x114;
        public const int WM_HSCROLLCLIPBOARD = 0x30E;
        public const int WM_ICONERASEBKGND = 0x27;
        public const int WM_IME_CHAR = 0x286;
        public const int WM_IME_COMPOSITION = 0x10F;
        public const int WM_IME_COMPOSITIONFULL = 0x284;
        public const int WM_IME_CONTROL = 0x283;
        public const int WM_IME_ENDCOMPOSITION = 0x10E;
        public const int WM_IME_KEYDOWN = 0x290;
        public const int WM_IME_KEYLAST = 0x10F;
        public const int WM_IME_KEYUP = 0x291;
        public const int WM_IME_NOTIFY = 0x282;
        public const int WM_IME_SELECT = 0x285;
        public const int WM_IME_SETCONTEXT = 0x281;
        public const int WM_IME_STARTCOMPOSITION = 0x10D;
        public const int WM_INITDIALOG = 0x110;
        public const int WM_INITMENU = 0x116;
        public const int WM_INITMENUPOPUP = 0x117;
        public const int WM_INPUTLANGCHANGE = 0x51;
        public const int WM_INPUTLANGCHANGEREQUEST = 0x50;
        public const int WM_KEYDOWN = 0x100;
        public const int WM_KEYFIRST = 0x100;
        public const int WM_KEYLAST = 0x108;
        public const int WM_KEYUP = 0x101;
        public const int WM_KILLFOCUS = 0x8;
        public const int WM_LBUTTONDBLCLK = 0x203;
        public const int WM_LBUTTONDOWN = 0x201;
        public const int WM_LBUTTONUP = 0x202;
        public const int WM_MBUTTONDBLCLK = 0x209;
        public const int WM_MBUTTONDOWN = 0x207;
        public const int WM_MBUTTONUP = 0x208;
        public const int WM_MDIACTIVATE = 0x222;
        public const int WM_MDICASCADE = 0x227;
        public const int WM_MDICREATE = 0x220;
        public const int WM_MDIDESTROY = 0x221;
        public const int WM_MDIGETACTIVE = 0x229;
        public const int WM_MDIICONARRANGE = 0x228;
        public const int WM_MDIMAXIMIZE = 0x225;
        public const int WM_MDINEXT = 0x224;
        public const int WM_MDIREFRESHMENU = 0x234;
        public const int WM_MDIRESTORE = 0x223;
        public const int WM_MDISETMENU = 0x230;
        public const int WM_MDITILE = 0x226;
        public const int WM_MEASUREITEM = 0x2C;
        public const int WM_MENUCHAR = 0x120;
        public const int WM_MENUSELECT = 0x11F;
        public const int WM_MOUSEACTIVATE = 0x21;
        public const int WM_MOUSEFIRST = 0x200;
        public const int WM_MOUSELAST = 0x209;
        public const int WM_MOUSEMOVE = 0x200;
        public const int WM_MOUSEWHEEL = 0x20A;
        public const int WM_MOUSEHWHEEL = 0x20E;
        public const int WM_MOVE = 0x3;
        public const int WM_MOVING = 0x216;
        public const int WM_NCACTIVATE = 0x86;
        public const int WM_NCCALCSIZE = 0x83;
        public const int WM_NCCREATE = 0x81;
        public const int WM_NCDESTROY = 0x82;
        public const int WM_NCHITTEST = 0x84;
        public const int WM_NCLBUTTONDBLCLK = 0xA3;
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int WM_NCLBUTTONUP = 0xA2;
        public const int WM_NCMBUTTONDBLCLK = 0xA9;
        public const int WM_NCMBUTTONDOWN = 0xA7;
        public const int WM_NCMBUTTONUP = 0xA8;
        public const int WM_NCMOUSEMOVE = 0xA0;
        public const int WM_NCPAINT = 0x85;
        public const int WM_NCRBUTTONDBLCLK = 0xA6;
        public const int WM_NCRBUTTONDOWN = 0xA4;
        public const int WM_NCRBUTTONUP = 0xA5;
        public const int WM_NEXTDLGCTL = 0x28;
        public const int WM_NEXTMENU = 0x213;
        public const int WM_NOTIFY = 0x4E;
        public const int WM_NOTIFYFORMAT = 0x55;
        public const int WM_NULL = 0x0;
        public const int WM_OTHERWINDOWCREATED = 0x42;
        public const int WM_OTHERWINDOWDESTROYED = 0x43;
        public const int WM_PAINT = 0xF;
        public const int WM_PAINTCLIPBOARD = 0x309;
        public const int WM_PAINTICON = 0x26;
        public const int WM_PALETTECHANGED = 0x311;
        public const int WM_PALETTEISCHANGING = 0x310;
        public const int WM_PARENTNOTIFY = 0x210;
        public const int WM_PASTE = 0x302;
        public const int WM_PENWINFIRST = 0x380;
        public const int WM_PENWINLAST = 0x38F;
        public const int WM_POWER = 0x48;
        public const int WM_POWERBROADCAST = 0x218;
        public const int WM_PRINT = 0x317;
        public const int WM_PRINTCLIENT = 0x318;
        public const int WM_PSD_ENVSTAMPRECT = WM_USER + 5;
        public const int WM_PSD_FULLPAGERECT = WM_USER + 1;
        public const int WM_PSD_GREEKTEXTRECT = WM_USER + 4;
        public const int WM_PSD_MARGINRECT = WM_USER + 3;
        public const int WM_PSD_MINMARGINRECT = WM_USER + 2;
        public const int WM_PSD_PAGESETUPDLG = WM_USER;
        public const int WM_PSD_YAFULLPAGERECT = WM_USER + 6;
        public const int WM_QUERYDRAGICON = 0x37;
        public const int WM_QUERYENDSESSION = 0x11;
        public const int WM_QUERYNEWPALETTE = 0x30F;
        public const int WM_QUERYOPEN = 0x13;
        public const int WM_QUEUESYNC = 0x23;
        public const int WM_QUIT = 0x12;
        public const int WM_RBUTTONDBLCLK = 0x206;
        public const int WM_RBUTTONDOWN = 0x204;
        public const int WM_RBUTTONUP = 0x205;
        public const int WM_RENDERALLFORMATS = 0x306;
        public const int WM_RENDERFORMAT = 0x305;
        public const int WM_SETCURSOR = 0x20;
        public const int WM_SETFOCUS = 0x7;
        public const int WM_SETFONT = 0x30;
        public const int WM_SETHOTKEY = 0x32;
        public const int WM_SETICON = 0x80;
        public const int WM_SETREDRAW = 0xB;
        public const int WM_SETTEXT = 0xC;
        public const int WM_SETTINGCHANGE = 0x1A;
        public const int WM_SHOWWINDOW = 0x18;
        public const int WM_SIZE = 0x5;
        public const int WM_SIZECLIPBOARD = 0x30B;
        public const int WM_SIZING = 0x214;
        public const int WM_SPOOLERSTATUS = 0x2A;
        public const int WM_STYLECHANGED = 0x7D;
        public const int WM_STYLECHANGING = 0x7C;
        public const int WM_SYSCHAR = 0x106;
        public const int WM_SYSCOLORCHANGE = 0x15;
        public const int WM_SYSCOMMAND = 0x112;
        public const int WM_SYSDEADCHAR = 0x107;
        public const int WM_SYSKEYDOWN = 0x104;
        public const int WM_SYSKEYUP = 0x105;
        public const int WM_TCARD = 0x52;
        public const int WM_TIMECHANGE = 0x1E;
        public const int WM_TIMER = 0x113;
        public const int WM_UNDO = 0x304;
        public const int WM_USER = 0x400;
        public const int WM_USERCHANGED = 0x54;
        public const int WM_VKEYTOITEM = 0x2E;
        public const int WM_VSCROLL = 0x115;
        public const int WM_VSCROLLCLIPBOARD = 0x30A;
        public const int WM_WINDOWPOSCHANGED = 0x47;
        public const int WM_WINDOWPOSCHANGING = 0x46;
        public const int WM_WININICHANGE = 0x1A;

        public const int WM_THEMECHANGED = 0x31A;

        public const int WM_ENTERSIZEMOVE = 0x231;
        public const int WM_EXITSIZEMOVE = 0x232;


        public const int WM_XBUTTONDOWN = 0x20B;
        public const int WM_XBUTTONUP = 0x20C;
        public const int WM_XBUTTONDBLCLK = 0x20D;

        public const int WM_NCXBUTTONDBLCLK = 0xAD;
        public const int WM_NCXBUTTONDOWN = 0xAB;
        public const int WM_NCXBUTTONUP = 0xAC;
        public const int WM_DPICHANGED = 0x2E0;
        public const int WM_DPICHANGED_BEFOREPARENT = 0x02E2;
        public const int WM_DPICHANGED_AFTERPARENT = 0x02E3;
        public const int WM_GETDPISCALEDSIZE = 0x02E4;

        //WM_XBUTTONUP
        // ReSharper restore InconsistentNaming


        public static int LoWord(int word) => NativeMethods.LoWord(word);

        public static int LoWord(IntPtr word) => NativeMethods.LoWord(word);

        public static short HiWord(IntPtr word) => NativeMethods.HiWord(word);
        public static short HiWord(int word) => NativeMethods.HiWord(word);


#if DEBUG_WIN_MSG
        public static string ToString(WindowMessage m)
        {
            return Msg(m.Msg, m.LParam);
        }

        public static string ToString(int msg)
        {
            return Msg(msg, IntPtr.Zero);
        }
        public static string Msg(WindowMessage m)
        {
            return Msg(m.Msg, m.LParam);
        }

        public static string Msg(int uMsg)
        {
            return Msg(uMsg, IntPtr.Zero);
        }

        [SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity"), SuppressMessage("Microsoft.Maintainability", "CA1505:AvoidUnmaintainableCode"),
         SuppressMessage("Microsoft.Globalization", "CA1305:SpecifyIFormatProvider", MessageId = "System.Int32.ToString(System.String)"),
         SuppressMessage("Microsoft.Globalization", "CA1305:SpecifyIFormatProvider", MessageId = "System.String.Format(System.String,System.Object,System.Object)"), DebuggerStepThrough()]
        public static string Msg(int uMsg, IntPtr LParam)
        {
            string sMsg = null;
            //const int WM_USER = 0x400;
            //const int WM_CAP_START = WM_USER;
            //const int WM_CAP_UNICODE_START = WM_CAP_START + 100;
            switch (uMsg)
            {
                case 0x204E:
                    sMsg = "WM_REDIRECTED_NOTIFY";
                    break;
                case 0x6:
                    sMsg = "WM_ACTIVATE";
                    break;
                case 0x1C:
                    sMsg = "WM_ACTIVATEAPP";
                    break;
                case WM_USER + 1104:
                    sMsg = "WM_ADSPROP_NOTIFY_APPLY";
                    break;
                case WM_USER + 1103:
                    sMsg = "WM_ADSPROP_NOTIFY_CHANGE";
                    break;
                case WM_USER + 1110:
                    sMsg = "WM_ADSPROP_NOTIFY_ERROR";
                    break;
                case WM_USER + 1107:
                    sMsg = "WM_ADSPROP_NOTIFY_EXIT";
                    break;
                case WM_USER + 1106:
                    sMsg = "WM_ADSPROP_NOTIFY_FOREGROUND";
                    break;
                case WM_USER + 1102:
                    sMsg = "WM_ADSPROP_NOTIFY_PAGEHWND";
                    break;
                case WM_USER + 1101:
                    sMsg = "WM_ADSPROP_NOTIFY_PAGEINIT";
                    break;
                case WM_USER + 1105:
                    sMsg = "WM_ADSPROP_NOTIFY_SETFOCUS";
                    break;
                case 0x360:
                    sMsg = "WM_AFXFIRST";
                    break;
                case 0x37F:
                    sMsg = "WM_AFXLAST";
                    break;
                case 0x8000:
                    sMsg = "WM_APP";
                    break;
                case 0x319:
                    sMsg = "WM_APPCOMMAND";
                    break;
                case 0x30C:
                    sMsg = "WM_ASKCBFORMATNAME";
                    break;
                case 0x4B:
                    sMsg = "WM_CANCELJOURNAL";
                    break;
                case 0x1F:
                    sMsg = "WM_CANCELMODE";
                    break;
                case 0x215:
                    sMsg = "WM_CAPTURECHANGED";
                    break;
                case WM_USER + 69:
                    sMsg = "WM_CAP_ABORT";
                    break;
                case WM_USER + 46:
                    sMsg = "WM_CAP_DLG_VIDEOCOMPRESSION";
                    break;
                case WM_USER + 43:
                    sMsg = "WM_CAP_DLG_VIDEODISPLAY";
                    break;
                case WM_USER + 41:
                    sMsg = "WM_CAP_DLG_VIDEOFORMAT";
                    break;
                case WM_USER + 42:
                    sMsg = "WM_CAP_DLG_VIDEOSOURCE";
                    break;
                case WM_USER + 10:
                    sMsg = "WM_CAP_DRIVER_CONNECT";
                    break;
                case WM_USER + 11:
                    sMsg = "WM_CAP_DRIVER_DISCONNECT";
                    break;
                case WM_USER + 14:
                    sMsg = "WM_CAP_DRIVER_CAPS";
                    break;
                case WM_USER + 12:
                    sMsg = "WM_CAP_DRIVER_NAMEA";
                    break;
                case WM_USER + 12 + 100:
                    sMsg = "WM_CAP_DRIVER_NAMEW | WM_CAP_DRIVER_NAME";
                    break;
                case WM_USER + 13 + 100:
                    sMsg = "WM_CAP_DRIVER_VERSIONW | WM_CAP_DRIVER_VERSION";
                    break;
                case WM_USER + 30:
                    sMsg = "WM_CAP_EDIT_COPY";
                    break;
                case WM_USER + 22:
                    sMsg = "WM_CAP_FILE_ALLOCATE";
                    break;
                case WM_USER + 21:
                    sMsg = "WM_CAP_FILE_CAPTURE_FILEA";
                    break;
                case WM_USER + 21 + 100:
                    sMsg = "WM_CAP_FILE_CAPTURE_FILE | WM_CAP_FILE_CAPTURE_FILEW";
                    break;
                case WM_USER + 23:
                    sMsg = "WM_CAP_FILE_SAVEASA";
                    break;
                case WM_USER + 23 + 100:
                    sMsg = "WM_CAP_FILE_SAVEAS | WM_CAP_FILE_SAVEASW";
                    break;
                case WM_USER + 25:
                    sMsg = "WM_CAP_FILE_SAVEDIBA";
                    break;
                case WM_USER + 25 + 100:
                    sMsg = "WM_CAP_FILE_SAVEDIB | WM_CAP_FILE_SAVEDIBW";
                    break;
                case WM_USER + 20:
                    sMsg = "WM_CAP_FILE_CAPTURE_FILEA";
                    break;
                case WM_USER + 20 + 100:
                    sMsg = "WM_CAP_FILE_CAPTURE_FILE | WM_CAP_FILE_CAPTURE_FILEW";
                    break;
                case WM_USER + 24:
                    sMsg = "WM_CAP_FILE_INFOCHUNK";
                    break;
                case WM_USER + 36:
                    sMsg = "WM_CAP_AUDIOFORMAT";
                    break;
                case WM_USER + 1:
                    sMsg = "WM_CAP_CAPSTREAMPTR | WM_CAP_CALLBACK_ERRORA | WM_CHOOSEFONT_GETLOGFONT | WM_PSD_FULLPAGERECT";
                    break;
                case WM_USER + 67:
                    sMsg = "WM_CAP_MCI_DEVICEA";
                    break;
                case WM_USER + 67 + 100:
                    sMsg = "WM_CAP_MCI_DEVICE | WM_CAP_MCI_DEVICEW";
                    break;
                case WM_USER + 65:
                    sMsg = "WM_CAP_SEQUENCE_SETUP";
                    break;
                case WM_USER + 54:
                    sMsg = "WM_CAP_STATUS";
                    break;
                case WM_USER + 8:
                    sMsg = "WM_CAP_USER_DATA";
                    break;
                case WM_USER + 44:
                    sMsg = "WM_CAP_VIDEOFORMAT";
                    break;
                case WM_USER + 60:
                    sMsg = "WM_CAP_GRAB_FRAME";
                    break;
                case WM_USER + 61:
                    sMsg = "WM_CAP_GRAB_FRAME_NOSTOP";
                    break;
                case WM_USER + 83:
                    sMsg = "WM_CAP_PAL_AUTOCREATE";
                    break;
                case WM_USER + 84:
                    sMsg = "WM_CAP_PAL_MANUALCREATE";
                    break;
                case WM_USER + 80:
                    sMsg = "WM_CAP_PAL_OPENA";
                    break;
                case WM_USER + 80 + 100:
                    sMsg = "WM_CAP_PAL_OPEN | WM_CAP_PAL_OPENW";
                    break;
                case WM_USER + 82:
                    sMsg = "WM_CAP_PAL_PASTE";
                    break;
                case WM_USER + 81:
                    sMsg = "WM_CAP_PAL_SAVEA";
                    break;
                case WM_USER + 62:
                    sMsg = "WM_CAP_SEQUENCE";
                    break;
                case WM_USER + 63:
                    sMsg = "WM_CAP_SEQUENCE_NOFILE";
                    break;
                case WM_USER + 35:
                    sMsg = "WM_CAP_AUDIOFORMAT";
                    break;
                case WM_USER + 85:
                    sMsg = "WM_CAP_CALLBACK_CAPCONTROL";
                    break;
                case WM_USER + 2 + 100:
                    sMsg = "WM_CAP_CALLBACK_ERROR | WM_CAP_CALLBACK_ERRORW | WM_CHOOSEFONT_SETFLAGS";
                    break;
                case WM_USER + 5:
                    sMsg = "WM_CAP_CALLBACK_FRAME | WM_PSD_ENVSTAMPRECT";
                    break;
                case WM_USER + 3:
                    sMsg = "WM_CAP_CALLBACK_STATUSA | WM_PSD_MARGINRECT";
                    break;
                case WM_USER + 3 + 100:
                    sMsg = "WM_CAP_CALLBACK_STATUS | WM_CAP_CALLBACK_STATUSW";
                    break;
                case WM_USER + 6:
                    sMsg = "WM_CAP_CALLBACK_VIDEOSTREAM | WM_PSD_YAFULLPAGERECT";
                    break;
                case WM_USER + 7:
                    sMsg = "WM_CAP_CALLBACK_WAVESTREAM";
                    break;
                case WM_USER + 4:
                    sMsg = "WM_CAP_CALLBACK_YIELD | WM_PSD_GREEKTEXTRECT";
                    break;
                case WM_USER + 66:
                    sMsg = "WM_CAP_MCI_DEVICEA";
                    break;
                case WM_USER + 66 + 100:
                    sMsg = "WM_CAP_MCI_DEVICE | WM_CAP_MCI_DEVICEW";
                    break;
                case WM_USER + 51:
                    sMsg = "WM_CAP_OVERLAY";
                    break;
                case WM_USER + 50:
                    sMsg = "WM_CAP_PREVIEW";
                    break;
                case WM_USER + 53:
                    sMsg = "WM_CAP_SCALE";
                    break;
                case WM_USER + 55:
                    sMsg = "WM_CAP_SCROLL";
                    break;
                case WM_USER + 64:
                    sMsg = "WM_CAP_SEQUENCE_SETUP";
                    break;
                case WM_USER + 9:
                    sMsg = "WM_CAP_USER_DATA";
                    break;
                case WM_USER + 45:
                    sMsg = "WM_CAP_VIDEOFORMAT";
                    break;
                case WM_USER + 72:
                    sMsg = "WM_CAP_SINGLE_FRAME";
                    break;
                case WM_USER + 71:
                    sMsg = "WM_CAP_SINGLE_FRAME_CLOSE";
                    break;
                case WM_USER + 70:
                    sMsg = "WM_CAP_SINGLE_FRAME_OPEN";
                    break;
                case WM_USER + 68:
                    sMsg = "WM_CAP_STOP";
                    break;
                case WM_USER + 81 + 100:
                    sMsg = "WM_CAP_END | WM_CAP_UNICODE_END | WM_CAP_PAL_SAVE | WM_CAP_PAL_SAVEW";
                    break;
                case WM_USER + 100:
                    sMsg = "WM_CAP_UNICODE_START";
                    break;
                case 0x30D:
                    sMsg = "WM_CHANGECBCHAIN";
                    break;
                case 0x127:
                    sMsg = "WM_CHANGEUISTATE";
                    break;
                case 0x102:
                    sMsg = "WM_CHAR";
                    break;
                case 0x2F:
                    sMsg = "WM_CHARTOITEM";
                    break;
                case 0x22:
                    sMsg = "WM_CHILDACTIVATE";
                    break;
                case WM_USER + 101:
                    sMsg = "WM_CHOOSEFONT_SETLOGFONT";
                    break;
                case 0x303:
                    sMsg = "WM_CLEAR";
                    break;
                case 0x10:
                    sMsg = "WM_CLOSE";
                    break;
                case 0x111:
                    sMsg = "WM_COMMAND";
                    break;
                case 0x44:
                    sMsg = "WM_COMMNOTIFY";
                    break;
                case 0x41:
                    sMsg = "WM_COMPACTING";
                    break;
                case 0x39:
                    sMsg = "WM_COMPAREITEM";
                    break;
                case 0x7B:
                    sMsg = "WM_CONTEXTMENU";
                    break;
                case 0x10A:
                    sMsg = "WM_CONVERTREQUEST";
                    break;
                case 0x10B:
                    sMsg = "WM_CONVERTRESULT";
                    break;
                case 0x301:
                    sMsg = "WM_COPY";
                    break;
                case 0x4A:
                    sMsg = "WM_COPYDATA";
                    break;
                case WM_USER + 1000:
                    sMsg = "WM_CPL_LAUNCH";
                    break;
                case WM_USER + 1001:
                    sMsg = "WM_CPL_LAUNCHED";
                    break;
                case 0x1:
                    sMsg = "WM_CREATE";
                    break;
                case 0x19:
                    sMsg = "WM_CTLCOLOR";
                    break;
                case 0x135:
                    sMsg = "WM_CTLCOLORBTN";
                    break;
                case 0x136:
                    sMsg = "WM_CTLCOLORDLG";
                    break;
                case 0x133:
                    sMsg = "WM_CTLCOLOREDIT";
                    break;
                case 0x134:
                    sMsg = "WM_CTLCOLORLISTBOX";
                    break;
                case 0x132:
                    sMsg = "WM_CTLCOLORMSGBOX";
                    break;
                case 0x137:
                    sMsg = "WM_CTLCOLORSCROLLBAR";
                    break;
                case 0x138:
                    sMsg = "WM_CTLCOLORSTATIC";
                    break;
                case 0x300:
                    sMsg = "WM_CUT";
                    break;
                case 0x3E0 + 4:
                    sMsg = "WM_DDE_ACK";
                    break;
                case 0x3E0 + 2:
                    sMsg = "WM_DDE_ADVISE";
                    break;
                case 0x3E0 + 5:
                    sMsg = "WM_DDE_DATA";
                    break;
                case 0x3E0 + 8:
                    sMsg = "WM_DDE_EXECUTE | WM_DDE_LAST";
                    break;
                case 0x3E0:
                    sMsg = "WM_DDE_FIRST | WM_DDE_INITIATE";
                    break;
                case 0x3E0 + 7:
                    sMsg = "WM_DDE_POKE";
                    break;
                case 0x3E0 + 6:
                    sMsg = "WM_DDE_REQUEST";
                    break;
                case 0x3E0 + 1:
                    sMsg = "WM_DDE_TERMINATE";
                    break;
                case 0x3E0 + 3:
                    sMsg = "WM_DDE_UNADVISE";
                    break;
                case 0x103:
                    sMsg = "WM_DEADCHAR";
                    break;
                case 0x2D:
                    sMsg = "WM_DELETEITEM";
                    break;
                case 0x2:
                    sMsg = "WM_DESTROY";
                    break;
                case 0x307:
                    sMsg = "WM_DESTROYCLIPBOARD";
                    break;
                case 0x219:
                    sMsg = "WM_DEVICECHANGE";
                    break;
                case 0x1B:
                    sMsg = "WM_DEVMODECHANGE";
                    break;
                case 0x7E:
                    sMsg = "WM_DISPLAYCHANGE";
                    break;
                case 0x308:
                    sMsg = "WM_DRAWCLIPBOARD";
                    break;
                case 0x2B:
                    sMsg = "WM_DRAWITEM";
                    break;
                case 0x233:
                    sMsg = "WM_DROPFILES";
                    break;
                case 0xA:
                    sMsg = "WM_ENABLE";
                    break;
                case 0x16:
                    sMsg = "WM_ENDSESSION";
                    break;
                case 0x121:
                    sMsg = "WM_ENTERIDLE";
                    break;
                case 0x211:
                    sMsg = "WM_ENTERMENULOOP";
                    break;
                case 0x231:
                    sMsg = "WM_ENTERSIZEMOVE";
                    break;
                case 0x14:
                    sMsg = "WM_ERASEBKGND";
                    break;
                case 0x212:
                    sMsg = "WM_EXITMENULOOP";
                    break;
                case 0x232:
                    sMsg = "WM_EXITSIZEMOVE";
                    break;
                case 0x1D:
                    sMsg = "WM_FONTCHANGE";
                    break;
                case 0x87:
                    sMsg = "WM_GETDLGCODE";
                    break;
                case 0x31:
                    sMsg = "WM_GETFONT";
                    break;
                case 0x33:
                    sMsg = "WM_GETHOTKEY";
                    break;
                case 0x7F:
                    sMsg = "WM_GETICON";
                    break;
                case 0x24:
                    sMsg = "WM_GETMINMAXINFO";
                    break;
                case 0x3D:
                    sMsg = "WM_GETOBJECT";
                    break;
                case 0xD:
                    sMsg = "WM_GETTEXT";
                    break;
                case 0xE:
                    sMsg = "WM_GETTEXTLENGTH";
                    break;
                case 0x358:
                    sMsg = "WM_HANDHELDFIRST";
                    break;
                case 0x35F:
                    sMsg = "WM_HANDHELDLAST";
                    break;
                case 0x53:
                    sMsg = "WM_HELP";
                    break;
                case 0x312:
                    sMsg = "WM_HOTKEY";
                    break;
                case 0x114:
                    sMsg = "WM_HSCROLL";
                    break;
                case 0x30E:
                    sMsg = "WM_HSCROLLCLIPBOARD";
                    break;
                case 0x27:
                    sMsg = "WM_ICONERASEBKGND";
                    break;
                case 0x286:
                    sMsg = "WM_IME_CHAR";
                    break;
                case 0x10F:
                    sMsg = "WM_IME_COMPOSITION | WM_IME_KEYLAST";
                    break;
                case 0x284:
                    sMsg = "WM_IME_COMPOSITIONFULL";
                    break;
                case 0x283:
                    sMsg = "WM_IME_CONTROL";
                    break;
                case 0x10E:
                    sMsg = "WM_IME_ENDCOMPOSITION";
                    break;
                case 0x290:
                    sMsg = "WM_IME_KEYDOWN";
                    break;
                case 0x291:
                    sMsg = "WM_IME_KEYUP";
                    break;
                case 0x282:
                    sMsg = "WM_IME_NOTIFY";
                    break;
                case 0x280:
                    sMsg = "WM_IME_REPORT";
                    break;
                case 0x288:
                    sMsg = "WM_IME_REQUEST";
                    break;
                case 0x285:
                    sMsg = "WM_IME_SELECT";
                    break;
                case 0x281:
                    sMsg = "WM_IME_SETCONTEXT";
                    break;
                case 0x10D:
                    sMsg = "WM_IME_STARTCOMPOSITION";
                    break;
                case 0x110:
                    sMsg = "WM_INITDIALOG";
                    break;
                case 0x116:
                    sMsg = "WM_INITMENU";
                    break;
                case 0x117:
                    sMsg = "WM_INITMENUPOPUP";
                    break;
                case 0xFF:
                    sMsg = "WM_INPUT";
                    break;
                case 0x51:
                    sMsg = "WM_INPUTLANGCHANGE";
                    break;
                case 0x50:
                    sMsg = "WM_INPUTLANGCHANGEREQUEST";
                    break;
                case 0x10C:
                    sMsg = "WM_INTERIM";
                    break;
                case 0x100:
                    sMsg = "WM_KEYDOWN | WM_KEYFIRST";
                    break;
                case 0x101:
                    sMsg = "WM_KEYUP";
                    break;
                case 0x8:
                    sMsg = "WM_KILLFOCUS";
                    break;
                case 0x203:
                    sMsg = "WM_LBUTTONDBLCLK";
                    break;
                case 0x201:
                    sMsg = "WM_LBUTTONDOWN";
                    break;
                case 0x202:
                    sMsg = "WM_LBUTTONUP";
                    break;
                case 0x209:
                    sMsg = "WM_MBUTTONDBLCLK";
                    break;
                case 0x207:
                    sMsg = "WM_MBUTTONDOWN";
                    break;
                case 0x208:
                    sMsg = "WM_MBUTTONUP";
                    break;
                case 0x222:
                    sMsg = "WM_MDIACTIVATE";
                    break;
                case 0x227:
                    sMsg = "WM_MDICASCADE";
                    break;
                case 0x220:
                    sMsg = "WM_MDICREATE";
                    break;
                case 0x221:
                    sMsg = "WM_MDIDESTROY";
                    break;
                case 0x229:
                    sMsg = "WM_MDIGETACTIVE";
                    break;
                case 0x228:
                    sMsg = "WM_MDIICONARRANGE";
                    break;
                case 0x225:
                    sMsg = "WM_MDIMAXIMIZE";
                    break;
                case 0x224:
                    sMsg = "WM_MDINEXT";
                    break;
                case 0x234:
                    sMsg = "WM_MDIREFRESHMENU";
                    break;
                case 0x223:
                    sMsg = "WM_MDIRESTORE";
                    break;
                case 0x230:
                    sMsg = "WM_MDISETMENU";
                    break;
                case 0x226:
                    sMsg = "WM_MDITILE";
                    break;
                case 0x2C:
                    sMsg = "WM_MEASUREITEM";
                    break;
                case 0x120:
                    sMsg = "WM_MENUCHAR";
                    break;
                case 0x126:
                    sMsg = "WM_MENUCOMMAND";
                    break;
                case 0x123:
                    sMsg = "WM_MENUDRAG";
                    break;
                case 0x124:
                    sMsg = "WM_MENUGETOBJECT";
                    break;
                case 0x122:
                    sMsg = "WM_MENURBUTTONUP";
                    break;
                case 0x11F:
                    sMsg = "WM_MENUSELECT";
                    break;
                case 0x21:
                    sMsg = "WM_MOUSEACTIVATE";
                    break;
                case 0x2A1:
                    sMsg = "WM_MOUSEHOVER";
                    break;
                case 0x2A3:
                    sMsg = "WM_MOUSELEAVE";
                    break;
                case 0x200:
                    sMsg = "WM_MOUSEMOVE";
                    break;
                case 0x20A:
                    sMsg = "WM_MOUSEWHEEL";
                    break;
                case 0x20E:
                    sMsg = "WM_MOUSEHWHEEL";
                    break;
                case 0x3:
                    sMsg = "WM_MOVE";
                    break;
                case 0x216:
                    sMsg = "WM_MOVING";
                    break;
                case 0x86:
                    sMsg = "WM_NCACTIVATE";
                    break;
                case 0x83:
                    sMsg = "WM_NCCALCSIZE";
                    break;
                case 0x81:
                    sMsg = "WM_NCCREATE";
                    break;
                case 0x82:
                    sMsg = "WM_NCDESTROY";
                    break;
                case 0x84:
                    sMsg = "WM_NCHITTEST";
                    break;
                case 0xA3:
                    sMsg = "WM_NCLBUTTONDBLCLK";
                    break;
                case 0xA1:
                    sMsg = "WM_NCLBUTTONDOWN";
                    break;
                case 0xA2:
                    sMsg = "WM_NCLBUTTONUP";
                    break;
                case 0xA9:
                    sMsg = "WM_NCMBUTTONDBLCLK";
                    break;
                case 0xA7:
                    sMsg = "WM_NCMBUTTONDOWN";
                    break;
                case 0xA8:
                    sMsg = "WM_NCMBUTTONUP";
                    break;
                case 0x2A0:
                    sMsg = "WM_NCMOUSEHOVER";
                    break;
                case 0x2A2:
                    sMsg = "WM_NCMOUSELEAVE";
                    break;
                case 0xA0:
                    sMsg = "WM_NCMOUSEMOVE";
                    break;
                case 0x85:
                    sMsg = "WM_NCPAINT";
                    break;
                case 0xA6:
                    sMsg = "WM_NCRBUTTONDBLCLK";
                    break;
                case 0xA4:
                    sMsg = "WM_NCRBUTTONDOWN";
                    break;
                case 0xA5:
                    sMsg = "WM_NCRBUTTONUP";
                    break;
                case 0xAD:
                    sMsg = "WM_NCXBUTTONDBLCLK";
                    break;
                case 0xAB:
                    sMsg = "WM_NCXBUTTONDOWN";
                    break;
                case 0xAC:
                    sMsg = "WM_NCXBUTTONUP";
                    break;
                case 0x28:
                    sMsg = "WM_NEXTDLGCTL";
                    break;
                case 0x213:
                    sMsg = "WM_NEXTMENU";
                    break;
                case 0x4E:
                    sMsg = "WM_NOTIFY";
                    break;
                case 0x55:
                    sMsg = "WM_NOTIFYFORMAT";
                    break;
                case 0x0:
                    sMsg = "WM_NULL";
                    break;
                case 0xF:
                    sMsg = "WM_PAINT";
                    break;
                case 0x309:
                    sMsg = "WM_PAINTCLIPBOARD";
                    break;
                case 0x26:
                    sMsg = "WM_PAINTICON";
                    break;
                case 0x311:
                    sMsg = "WM_PALETTECHANGED";
                    break;
                case 0x310:
                    sMsg = "WM_PALETTEISCHANGING";
                    break;
                case 0x210:
                    sMsg = "WM_PARENTNOTIFY";
                    break;
                case 0x302:
                    sMsg = "WM_PASTE";
                    break;
                case 0x380:
                    sMsg = "WM_PENWINFIRST";
                    break;
                case 0x38F:
                    sMsg = "WM_PENWINLAST";
                    break;
                case 0x48:
                    sMsg = "WM_POWER";
                    break;
                case 0x218:
                    sMsg = "WM_POWERBROADCAST";
                    break;
                case 0x317:
                    sMsg = "WM_PRINT";
                    break;
                case 0x318:
                    sMsg = "WM_PRINTCLIENT";
                    break;
                case WM_USER + 2:
                    sMsg = "WM_PSD_MINMARGINRECT";
                    break;
                case 0x37:
                    sMsg = "WM_QUERYDRAGICON";
                    break;
                case 0x11:
                    sMsg = "WM_QUERYENDSESSION";
                    break;
                case 0x30F:
                    sMsg = "WM_QUERYNEWPALETTE";
                    break;
                case 0x13:
                    sMsg = "WM_QUERYOPEN";
                    break;
                case 0x129:
                    sMsg = "WM_QUERYUISTATE";
                    break;
                case 0x23:
                    sMsg = "WM_QUEUESYNC";
                    break;
                case 0x12:
                    sMsg = "WM_QUIT";
                    break;
                case 0xCCCD:
                    sMsg = "WM_RASDIALEVENT";
                    break;
                case 0x206:
                    sMsg = "WM_RBUTTONDBLCLK";
                    break;
                case 0x204:
                    sMsg = "WM_RBUTTONDOWN";
                    break;
                case 0x205:
                    sMsg = "WM_RBUTTONUP";
                    break;
                case 0x306:
                    sMsg = "WM_RENDERALLFORMATS";
                    break;
                case 0x305:
                    sMsg = "WM_RENDERFORMAT";
                    break;
                case 0x20:
                    sMsg = "WM_SETCURSOR";
                    break;
                case 0x7:
                    sMsg = "WM_SETFOCUS";
                    break;
                case 0x30:
                    sMsg = "WM_SETFONT";
                    break;
                case 0x32:
                    sMsg = "WM_SETHOTKEY";
                    break;
                case 0x80:
                    sMsg = "WM_SETICON";
                    break;
                case 0xB:
                    sMsg = "WM_SETREDRAW";
                    break;
                case 0xC:
                    sMsg = "WM_SETTEXT";
                    break;
                case 0x1A:
                    sMsg = "WM_SETTINGCHANGE";
                    break;
                case 0x18:
                    sMsg = "WM_SHOWWINDOW";
                    break;
                case 0x5:
                    sMsg = "WM_SIZE";
                    break;
                case 0x30B:
                    sMsg = "WM_SIZECLIPBOARD";
                    break;
                case 0x214:
                    sMsg = "WM_SIZING";
                    break;
                case 0x2A:
                    sMsg = "WM_SPOOLERSTATUS";
                    break;
                case 0x7D:
                    sMsg = "WM_STYLECHANGED";
                    break;
                case 0x7C:
                    sMsg = "WM_STYLECHANGING";
                    break;
                case 0x88:
                    sMsg = "WM_SYNCPAINT";
                    break;
                case 0x106:
                    sMsg = "WM_SYSCHAR";
                    break;
                case 0x15:
                    sMsg = "WM_SYSCOLORCHANGE";
                    break;
                case 0x112:
                    sMsg = "WM_SYSCOMMAND";
                    break;
                case 0x107:
                    sMsg = "WM_SYSDEADCHAR";
                    break;
                case 0x104:
                    sMsg = "WM_SYSKEYDOWN";
                    break;
                case 0x105:
                    sMsg = "WM_SYSKEYUP";
                    break;
                case 0x2C0:
                    sMsg = "WM_TABLET_FIRST";
                    break;
                case 0x2DF:
                    sMsg = "WM_TABLET_LAST";
                    break;
                case 0x52:
                    sMsg = "WM_TCARD";
                    break;
                case 0x31A:
                    sMsg = "WM_THEMECHANGED";
                    break;
                case 0x1E:
                    sMsg = "WM_TIMECHANGE";
                    break;
                case 0x113:
                    sMsg = "WM_TIMER";
                    break;
                case 0x304:
                    sMsg = "WM_UNDO";
                    break;
                case 0x109:
                    sMsg = "WM_KEYLAST | WM_UNICHAR | WM_WNT_CONVERTREQUESTEX";
                    break;
                case 0x125:
                    sMsg = "WM_UNINITMENUPOPUP";
                    break;
                case 0x128:
                    sMsg = "WM_UPDATEUISTATE";
                    break;
                case WM_USER:
                    sMsg = "WM_USER | WM_PSD_PAGESETUPDLG | WM_CAP_PREVIEWRATE | WM_CAP_DRIVER_VERSIONA | WM_CAP_START";
                    break;
                case 0x54:
                    sMsg = "WM_USERCHANGED";
                    break;
                case 0x2E:
                    sMsg = "WM_VKEYTOITEM";
                    break;
                case 0x115:
                    sMsg = "WM_VSCROLL";
                    break;
                case 0x30A:
                    sMsg = "WM_VSCROLLCLIPBOARD";
                    break;
                case 0x47:
                    sMsg = "WM_WINDOWPOSCHANGED";
                    break;
                case 0x46:
                    sMsg = "WM_WINDOWPOSCHANGING";
                    break;
                case 0x2B1:
                    sMsg = "WM_WTSSESSION_CHANGE";
                    break;
                case 0x20D:
                    sMsg = "WM_XBUTTONDBLCLK";
                    break;
                case 0x20B:
                    sMsg = "WM_XBUTTONDOWN";
                    break;
                case 0x20C:
                    sMsg = "WM_XBUTTONUP";
                    break;
 
                case WM_DPICHANGED:
                    sMsg = "WM_DPICHANGED";
                    break;
                case WM_DPICHANGED_BEFOREPARENT:
                    sMsg = "WM_DPICHANGED_BEFOREPARENT";
                    break;
                case WM_DPICHANGED_AFTERPARENT:
                    sMsg = "WM_DPICHANGED_AFTERPARENT";
                    break;
                case WM_GETDPISCALEDSIZE:
                    sMsg = "WM_GETDPISCALEDSIZE";
                    break;
                default:
                    if (uMsg < WM_USER)
                    {
                        sMsg = "0x" + uMsg.ToString("X") + " (unknown WM_ msg)";
                    }
                    else if (uMsg >= WM_USER && uMsg <= 0x7FFF)
                    {
                        sMsg = "WM_USER + " + (uMsg - WM_USER) + " (control specific msg)";
                        //    Case WM_USER To &H7FFF: sMsg = GetTVMsgStr(uMsg)
                    }
                    else if (uMsg >= 0x8000L && uMsg <= 0xBFFFL)
                    {
                        sMsg = "0x" + uMsg.ToString("X") + " (reserved msg)";
                    }
                    else if (uMsg >= 0xC000L && uMsg <= 0xFFFFL)
                    {
                        sMsg = "0x" + uMsg.ToString("X") + " (registered msg)";
                    }
                    else
                    {
                        sMsg = "0x" + uMsg.ToString("X") + " (reserved msg)";
                    }
                    break;
            }
            if (uMsg == 0x204E || uMsg == 0x4E)
            {
                sMsg = sMsg + $", From:{Marshal.ReadInt32(LParam):X}, Code:{Marshal.ReadInt32(LParam, 8)}";
            }
            return sMsg;
        }



        [DebuggerStepThrough()]
        public static string WindowPosFlags(int flags)
        {
            string flagNames = null;

            AddFlag(0x1, "SWP_NOSIZE", ref flags, ref flagNames);
            AddFlag(0x2, "SWP_NOMOVE", ref flags, ref flagNames);
            AddFlag(0x4, "SWP_NOZORDER", ref flags, ref flagNames);
            AddFlag(0x8, "SWP_NOREDRAW", ref flags, ref flagNames);
            AddFlag(0x10, "SWP_NOACTIVATE", ref flags, ref flagNames);
            AddFlag(0x20, "SWP_FRAMECHANGED", ref flags, ref flagNames); //| SWP_DRAWFRAME
            AddFlag(0x40, "SWP_SHOWWINDOW", ref flags, ref flagNames);
            AddFlag(0x80, "SWP_HIDEWINDOW", ref flags, ref flagNames);
            AddFlag(0x100, "SWP_NOCOPYBITS", ref flags, ref flagNames);
            AddFlag(0x200, "SWP_NOOWNERZORDER", ref flags, ref flagNames);
            AddFlag(0x2000, "SWP_DEFERERASE", ref flags, ref flagNames);
            AddFlag(0x4000, "SWP_NOOWNERZORDER", ref flags, ref flagNames); //| SWP_NOREPOSITION

            TrimFlags(flags, ref flagNames);

            return flagNames;
        }



        [DebuggerStepThrough()]
        public static string ClassStyles(int flags)
        {
            return ClassStyles(flags, " | ");
        }

        [DebuggerStepThrough()]
        public static string ClassStyles(int flags, string delimiter)
        {
            string flagNames = null;

            AddFlag(0x1, "CS_VREDRAW", ref flags, ref flagNames, delimiter);
            AddFlag(0x2, "CS_HREDRAW", ref flags, ref flagNames, delimiter);
            AddFlag(0x8, "CS_DBLCLKS", ref flags, ref flagNames, delimiter);
            AddFlag(0x20, "CS_OWNDC", ref flags, ref flagNames, delimiter);
            AddFlag(0x40, "CS_CLASSDC", ref flags, ref flagNames, delimiter);
            AddFlag(0x80, "CS_PARENTDC", ref flags, ref flagNames, delimiter); //| SWP_DRAWFRAME
            AddFlag(0x200, "CS_NOCLOSE", ref flags, ref flagNames, delimiter);
            AddFlag(0x800, "CS_SAVEBITS", ref flags, ref flagNames, delimiter);
            AddFlag(0x1000, "CS_BYTEALIGNCLIENT", ref flags, ref flagNames, delimiter);
            AddFlag(0x2000, "CS_BYTEALIGNWINDOW", ref flags, ref flagNames, delimiter);
            AddFlag(0x4000, "CS_GLOBALCLASS", ref flags, ref flagNames, delimiter);

            AddFlag(0x10000, "CS_IME", ref flags, ref flagNames, delimiter);
            AddFlag(0x20000, "CS_DROPSHADOW", ref flags, ref flagNames, delimiter);

            TrimFlags(flags, ref flagNames, delimiter);

            return flagNames;
        }



        [DebuggerStepThrough()]
        public static string StyleFlags(int flags)
        {
            return StyleFlags(flags, " | ");
        }

        public static string StyleFlags(int flags, string delimiter)
        {
            string flagNames = null;
            AddFlag(0x0, "WS_OVERLAPPED", ref flags, ref flagNames, delimiter); //WS_SYSMENU
            AddFlag(0x80000, "WS_SYSMENU", ref flags, ref flagNames, delimiter); //WS_SYSMENU
            AddFlag(0x10000000, "WS_VISIBLE", ref flags, ref flagNames, delimiter); //WS_VISIBLE
            AddFlag(0x800000, "WS_BORDER", ref flags, ref flagNames, delimiter); //WS_BORDER
            AddFlag(0xC00000, "WS_CAPTION", ref flags, ref flagNames, delimiter); //WS_CAPTION
            AddFlag(0x40000000, "WS_CHILD", ref flags, ref flagNames, delimiter); //WS_CHILD/WS_CHILDWINDOW
            AddFlag(0x2000000, "WS_CLIPCHILDREN", ref flags, ref flagNames, delimiter); //WS_CLIPCHILDREN
            AddFlag(0x4000000, "WS_CLIPSIBLINGS", ref flags, ref flagNames, delimiter); //WS_CLIPSIBLINGS
            AddFlag(0x8000000, "WS_DISABLED", ref flags, ref flagNames, delimiter); //WS_DISABLED
            AddFlag(0x400000, "WS_DLGFRAME", ref flags, ref flagNames, delimiter); //WS_DLGFRAME
            AddFlag(0x20000, "WS_GROUP", ref flags, ref flagNames, delimiter); //WS_GROUP
            AddFlag(0x100000, "WS_HSCROLL", ref flags, ref flagNames, delimiter); //WS_HSCROLL
            AddFlag(0x20000000, "WS_MINIMIZE", ref flags, ref flagNames, delimiter); //WS_MINIMIZE/WS_ICONIC
            AddFlag(0x1000000, "WS_MAXIMIZE", ref flags, ref flagNames, delimiter); //WS_MAXIMIZE
            AddFlag(0x20000, "WS_MINIMIZEBOX", ref flags, ref flagNames, delimiter); //WS_MINIMIZEBOX
            AddFlag(0x10000, "WS_MAXIMIZEBOX", ref flags, ref flagNames, delimiter); //WS_MAXIMIZEBOX
            //Call AddFlag(&H80000000, "WS_OVERLAPPED", flags, flagNames)    'WS_OVERLAPPED
            AddFlag(unchecked((int) 0x80000000), "WS_POPUP", ref flags, ref flagNames, delimiter); //WS_POPUP
            AddFlag(0x40000, "WS_THICKFRAME", ref flags, ref flagNames, delimiter); //WS_THICKFRAME/WS_SIZEBOX
            AddFlag(0x10000, "WS_TABSTOP", ref flags, ref flagNames, delimiter); //WS_TABSTOP
            AddFlag(0x200000, "WS_VSCROLL", ref flags, ref flagNames, delimiter); //WS_VSCROLL

            TrimFlags(flags, ref flagNames, delimiter);

            return flagNames;
        }



        [DebuggerStepThrough()]
        public static string ExStyleFlags(int flags)
        {
            return ExStyleFlags(flags, " | ");
        }

        public static string ExStyleFlags(int flags, string delimiter)
        {
            string flagNames = string.Format(System.Globalization.CultureInfo.InvariantCulture, "WS_EX_RIGHTSCROLLBAR{0}WS_EX_LEFT{0}WS_EX_LTRREADING{0}", " | ");


            AddFlag(0x1, "WS_EX_DLGMODALFRAME", ref flags, ref flagNames, delimiter); // WS_EX_DLGMODALFRAME
            AddFlag(0x4, "WS_EX_NOPARENTNOTIFY", ref flags, ref flagNames, delimiter); // WS_EX_NOPARENTNOTIFY
            AddFlag(0x8, "WS_EX_TOPMOST", ref flags, ref flagNames, delimiter); // WS_EX_TOPMOST
            AddFlag(0x10, "WS_EX_ACCEPTFILES", ref flags, ref flagNames, delimiter); // WS_EX_ACCEPTFILES
            AddFlag(0x20, "WS_EX_TRANSPARENT", ref flags, ref flagNames, delimiter); // WS_EX_TRANSPARENT
            AddFlag(0x40, "WS_EX_MDICHILD", ref flags, ref flagNames, delimiter); // WS_EX_MDICHILD
            AddFlag(0x80, "WS_EX_TOOLWINDOW", ref flags, ref flagNames, delimiter); // WS_EX_TOOLWINDOW
            AddFlag(0x100, "WS_EX_WINDOWEDGE", ref flags, ref flagNames, delimiter); // WS_EX_WINDOWEDGE
            AddFlag(0x200, "WS_EX_CLIENTEDGE", ref flags, ref flagNames, delimiter); // WS_EX_CLIENTEDGE
            AddFlag(0x400, "WS_EX_CONTEXTHELP", ref flags, ref flagNames, delimiter); // WS_EX_CONTEXTHELP
            AddFlag(0x1000, "WS_EX_RIGHT", ref flags, ref flagNames, delimiter); // WS_EX_RIGHT
            AddFlag(0x2000, "WS_EX_RTLREADING", ref flags, ref flagNames, delimiter); // WS_EX_RTLREADING
            AddFlag(0x4000, "WS_EX_LEFTSCROLLBAR", ref flags, ref flagNames, delimiter); // WS_EX_LEFTSCROLLBAR
            AddFlag(0x10000, "WS_EX_CONTROLPARENT", ref flags, ref flagNames, delimiter); // WS_EX_CONTROLPARENT
            AddFlag(0x20000, "WS_EX_STATICEDGE", ref flags, ref flagNames, delimiter); // WS_EX_STATICEDGE
            AddFlag(0x40000, "WS_EX_APPWINDOW", ref flags, ref flagNames, delimiter); // WS_EX_APPWINDOW

            AddFlag(0x80000, "WS_EX_LAYERED", ref flags, ref flagNames, delimiter); // WS_EX_LAYERED
            AddFlag(0x100000, "WS_EX_NOINHERITLAYOUT", ref flags, ref flagNames, delimiter); // WS_EX_NOINHERITLAYOUT - Disable inheritence of mirroring by children
            AddFlag(0x100000, "WS_EX_LAYOUTRTL", ref flags, ref flagNames, delimiter); // WS_EX_LAYOUTRTL - Right to left mirroring
            AddFlag(0x8000000, "WS_EX_NOACTIVATE", ref flags, ref flagNames, delimiter); // WS_EX_NOACTIVATE
            AddFlag(0x2000000, "WS_EX_COMPOSITED", ref flags, ref flagNames, delimiter); // WS_EX_COMPOSITED

            TrimFlags(flags, ref flagNames, delimiter);

            return flagNames;
        }



        [DebuggerStepThrough()]
        public static string ScrollBarFlags(int message, IntPtr wParam)
        {
            const int WM_HSCROLL = 0x114;
            const int WM_VSCROLL = 0x115;

            string scrollInfo = string.Empty;
            int loWordHiWord = NativeMethods.IntPtrToInt32(wParam);
            int lWord = NativeMethods.LoWord(loWordHiWord);

            if (message == WM_HSCROLL || message == WM_VSCROLL)
            {
                var isHScroll = message == WM_HSCROLL;

                scrollInfo = (isHScroll ? "WM_HSCROLL,(" : "WM_VSCROLL, (") + lWord.ToString() + ") ";
                switch (lWord)
                {
                    case 0:
                        scrollInfo += isHScroll ? "SB_LINELEFT" : "SB_LINEUP";
                        break;
                    case 1:
                        scrollInfo += isHScroll ? "SB_LINERIGHT" : "SB_LINEDOWN";
                        break;
                    case 2:
                        scrollInfo += isHScroll ? "SB_PAGELEFT " : "SB_PAGEUP";
                        break;
                    case 3:
                        scrollInfo += isHScroll ? "SB_PAGERIGHT" : "SB_PAGEDOWN";
                        break;
                    case 4:
                        scrollInfo += "SB_THUMBPOSITION " + NativeMethods.HiWord(loWordHiWord).ToString();
                        break;
                    case 5:
                        scrollInfo += "SB_THUMBTRACK " + NativeMethods.HiWord(loWordHiWord).ToString();
                        break;
                    case 6:
                        scrollInfo += isHScroll ? "SB_LEFT" : "SB_TOP";
                        break;
                    case 7:
                        scrollInfo += isHScroll ? "SB_RIGHT" : "SB_BOTTOM";
                        break;
                    case 8:
                        scrollInfo += "SB_ENDSCROLL";
                        break;
                    default:
                        scrollInfo += lWord.ToString();
                        break;
                }
            }

            return scrollInfo;
        }



        [DebuggerStepThrough()]
        public static string WmHitTestCode(int code)
        {
            switch (code)
            {
                case -2:
                    return "HTERROR";
                case -1:
                    return "HTTRANSPARENT";
                case 0:
                    return "HTNOWHERE";
                case 1:
                    return "HTCLIENT";
                case 2:
                    return "HTCAPTION";
                case 3:
                    return "HTSYSMENU";
                case 4: //HTSIZE
                    return "HTGROWBOX";
                case 5:
                    return "HTMENU";
                case 6:
                    return "HTHSCROLL";
                case 7:
                    return "HTVSCROLL";
                case 8: //HTREDUCE
                    return "HTMINBUTTON";
                case 9: //HTZOOM
                    return "HTMAXBUTTON";
                case 10: //HTSIZEFIRST
                    return "HTLEFT";
                case 11:
                    return "HTRIGHT";
                case 12:
                    return "HTTOP";
                case 13:
                    return "HTTOPLEFT";
                case 14:
                    return "HTTOPRIGHT";
                case 15:
                    return "HTBOTTOM";
                case 16:
                    return "HTBOTTOMLEFT";
                case 17: //HTSIZELAST
                    return "HTBOTTOMRIGHT";
                case 18:
                    return "HTBORDER";
                //#if(WINVER >= 00400)
                case 19:
                    return "HTOBJECT";
                case 20:
                    return "HTCLOSE";
                case 21:
                    return "HTHELP";
                //#endif /* WINVER >= 0x0400 */
                default:
                    return null;
            }
        }




        //<System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1305:SpecifyIFormatProvider", MessageId:="System.Int32.ToString(System.String)")>
        //<DebuggerStepThrough()>
        private static void TrimFlags(int flags, ref string flagNames, string delimiter = " or ")
        {
            if (flags != 0)
            {
                flagNames = flagNames + flags.ToString("X");
            }
            else if (!string.IsNullOrEmpty(flagNames))
            {
                flagNames = flagNames.Substring(0, flagNames.Length - delimiter.Length);
            }
        }



        [DebuggerStepThrough()]
        private static void AddFlag(int flag, string flagName, ref int flags, ref string flagNames, string delimiter = " | ")
        {
            if ((flags & flag) == flag)
            {
                flagNames = flagNames + flagName + delimiter;
                flags &= ~flag;
            }
        }


#endif
    }
}