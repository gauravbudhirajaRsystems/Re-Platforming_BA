//*************************************************
//* © 2020 Litera Corp. All Rights Reserved.
//**************************************************

using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using LevitJames.Core;
using LevitJames.Core.Diagnostics;
using LevitJames.Libraries.Hooks;
using LevitJames.Win32;
using Microsoft.Office.Interop.Word;
using static LevitJames.Libraries.NativeMethods;

// ReSharper disable once CheckNamespace
namespace LevitJames.MSOffice.MSWord
{
    //OpusApp
    //  _WwF
    //     _WwB
    //        _WwG
    //     _WwC
    //     NUIScrollbar

    #region WordDocumentEventArgs (class) 

    [Serializable]
    public class WordDocumentEventArgs : EventArgs
    {
        internal WordDocumentEventArgs(Document document)
        {
            Document = document;
        }

        /// <summary>
        ///     Returns a Word Window instance.
        /// </summary>
        /// <returns>A valid Window instance.</returns>
        public Document Document { get; }
    }

    #endregion // WordDocumentEventArgs (class)
 
    #region WordWindowActivateEventArgs (class)

    [Serializable]
    public sealed class WordWindowActivateEventArgs : WordWindowEventArgs
    {
	    internal WordWindowActivateEventArgs(Window window, IntPtr handle, bool fromNonClientArea, bool isActive) : base(window, handle)
	    {
		    FromNonClientArea = fromNonClientArea;
		    IsActive = isActive;
	    }

	    public bool FromNonClientArea { get; }

	    public bool IsActive { get; }
    }

    #endregion // WordWindowActivateEventArgs (class)

    #region WordWindowEventArgs (class)

    [Serializable]
    public class WordWindowEventArgs : EventArgs
    {
        internal WordWindowEventArgs(Window window)
        {
            Window = window;
        }

        internal WordWindowEventArgs(Window window, IntPtr handle)
        {
            Window = window;
            Handle = handle;
        }

        /// <summary>
        ///     The Window handle of the Window
        /// </summary>
        /// <returns>An IntPtr representing the Window's _WwG window.</returns>
        internal IntPtr Handle { get; }

        /// <summary>
        ///     Returns a Word Window instance.
        /// </summary>
        /// <returns>A valid Window instance.</returns>
        public Window Window { get; }


    }

    #endregion // WordWindowEventArgs (class)

    #region WordWindowStateChangedEventArgs (class) 

    [Serializable]
    public sealed class WordWindowStateChangedEventArgs : WordWindowEventArgs
    {
        public WordWindowStateChangedEventArgs(Window window, IntPtr handle, int windowState)
            : base(window, handle)
        {
            WindowState = windowState;
        }

        public int WindowState { get; }
    }

    #endregion // WordWindowStateChangedEventArgs (class)

    #region FocusChangedEventArgs (class) 

    [Serializable]
    public sealed class FocusChangedEventArgs : EventArgs
    {
        public FocusChangedEventArgs(IntPtr oldFocus, IntPtr newFocus)
        {
            OldFocus = oldFocus;
            NewFocus = newFocus;
        }
        public IntPtr OldFocus { get; }

        public IntPtr NewFocus { get; }
    }

    #endregion // FocusChangedEventArgs (class)

    //<DebuggerStepThrough()>
    public sealed class WordWindowHelper : IDisposable
    {

        //private int _tempState; // For debugging

        private IntPtr _hWndWindow;
        private int _inGetWindow;
        private Atom _classAtom;
        private OfficeUITheme _theme;
        private WordBoolean _windowsInTaskBar;
        private bool _hooked;

        private static Atom _msoWorkPaneClassAtom;
        private static Atom _opusAppClassAtom;
        private static Atom _wwGClassAtom;
        private static Atom _wwBClassAtom;
        private static Atom _wwFClassAtom;
 
        private Application _wordApplication;
   
        public event EventHandler WindowsInTaskBarChanging;
        public event EventHandler WindowsInTaskBarChanged;
        public event EventHandler<WordWindowStateChangedEventArgs> WindowStateChanged;
        public event EventHandler<WordWindowActivateEventArgs> WindowActivateChanged;

        public event EventHandler<WordWindowEventArgs> WindowClosed;
        public event EventHandler WordThemeChanged;
        public event EventHandler<FocusChangedEventArgs> FocusChanged;
        public event EventHandler<FocusChangedEventArgs> FocusLost;
        //public event EventHandler<WordDocumentEventArgs> DocumentClosed;
#if (TRACK_DISPOSED)
	    private readonly string _disposedSource;
#endif
         
        private WordWindowHelper(Application wordApplication)
        {
            _wordApplication = wordApplication;
#if (TRACK_DISPOSED)
		    _disposedSource = Core.Diagnostics.DisposeTracker.GetSource();
#endif
        }

        public static WordWindowHelper NewWordWindowHelper(Application wordApplication)
        {
            Check.NotNull(wordApplication, nameof(wordApplication));

            var wordWindowHelper = new WordWindowHelper(wordApplication) {_windowsInTaskBar = WordExtensions.WordApplication.ShowWindowsInTaskBarLJ()};

            AppDiagnostics.OptionChanged += wordWindowHelper.AppDiagnosticsChangedHandler;

            if (AppDiagnostics.GetOption(AppDiagnosticOptions.NoHooks) == false)
            {
                wordWindowHelper.SetHooks(hook: true);
            }

            return wordWindowHelper;
        }


        private void AppDiagnosticsChangedHandler(object sender, AppDiagnosticOptionsChangedEventArgs e)
        {
            if (e.HasItemChanged(AppDiagnosticOptions.NoHooks))
            {
                SetHooks(AppDiagnostics.GetOption(AppDiagnosticOptions.NoHooks) == false);
            }
        }


        internal static Atom OpusAppWindowClass
        {
            get
            {
                if (_opusAppClassAtom == 0)
                    _opusAppClassAtom = GetClassAtomFromName("OpusApp");
                return _opusAppClassAtom;
            }
        }


        internal static Atom MsoWorkPaneWindowClass
        {
            get
            {
                if (_msoWorkPaneClassAtom == 0)
                    _msoWorkPaneClassAtom = GetClassAtomFromName("MsoWorkPane");
                return _msoWorkPaneClassAtom;
            }
        }

        internal static Atom WwGWindowClass
        {
            [DebuggerStepThrough]
            get
            {
                if (_wwGClassAtom == 0)
                    _wwGClassAtom = GetClassAtomFromName("_WwG");

                return _wwGClassAtom;
            }
        }


        internal static Atom WwBWindowClass
        {
            get
            {
                if (_wwBClassAtom == 0)
                    _wwBClassAtom = GetClassAtomFromName("_WwB");

                return _wwBClassAtom;
            }
        }


        internal static Atom WwFWindowClass
        {
            get
            {
                if (_wwFClassAtom == 0)
                    _wwFClassAtom = GetClassAtomFromName("_WwF");

                return _wwFClassAtom;
            }
        }

         

        public OfficeUITheme Theme
        {
            get
            {
                if (_theme == OfficeUITheme.None)
                    _theme = _wordApplication.GetOfficeTheme();

                return _theme;
            }
        }


        public bool IsWordApplicationEnabled()
        {
            var handle = GetApplicationHandle();
            var ww = new WindowInfo(handle);
            return ww.Enabled;
        }


        internal static bool IsWindowClass(IntPtr handle, Atom classAtom)
        {
            return GetClassAtom(handle) == classAtom;
        }



        public static Document GetActiveDocument()
        {
            var docHandle = WindowInfo.FromActiveWindow();
            if (!docHandle.IsValid)
                return null;

            Window window = null;
            return TryGetWindowFromHandle(docHandle.Handle, ref window) ? window.Document : null;
        }


        public static IntPtr NormalizeHandleToWwG(IntPtr handle)
        {
            var hWndNormalized = default(IntPtr);

            var classAtom = GetClassAtom(handle);

            if (GetClassAtom(handle) == WwGWindowClass)
            {
                hWndNormalized = handle;
            }
            else if (classAtom == WwBWindowClass)
            {
                hWndNormalized = FindWindowEx(handle, IntPtr.Zero, WwGWindowClass, lpszWindow: null);
            }
            else if (classAtom == WwFWindowClass)
            {
                handle = FindWindowEx(handle, IntPtr.Zero, WwBWindowClass, lpszWindow: null);
                if (handle != IntPtr.Zero)
                {
                    hWndNormalized = FindWindowEx(handle, IntPtr.Zero, WwGWindowClass, lpszWindow: null);
                }
            }
            else if (classAtom == OpusAppWindowClass)
            {
                handle = FindWindowEx(handle, IntPtr.Zero, WwFWindowClass, lpszWindow: null);
                if (handle != IntPtr.Zero)
                {
                    handle = FindWindowEx(handle, IntPtr.Zero, WwBWindowClass, lpszWindow: null);
                    if (handle != IntPtr.Zero)
                    {
                        hWndNormalized = FindWindowEx(handle, IntPtr.Zero, WwGWindowClass, lpszWindow: null);
                    }
                }
            }

            //Other windows may also be _Wwg Class windows.

            while (!IsWwGDocumentWindow(hWndNormalized))
            {
                hWndNormalized = FindWindowEx(handle, hWndNormalized, WwGWindowClass, lpszWindow: null);
                if (IsWindow(hWndNormalized) == false)
                {
                    break;
                }
            }

            if (IsWindow(hWndNormalized) == false)
            {
                hWndNormalized = IntPtr.Zero;
            }

            return hWndNormalized;
        }


        public IntPtr GetApplicationHandle()
        {
            return GetApplicationHandleInternal(window: null);
        }

        public IntPtr GetApplicationHandle(Window window)
        {
            return GetApplicationHandleInternal(window);
        }

        private IntPtr GetApplicationHandleInternal(Window window)
        {
            var hWnd = default(IntPtr);
            var classAtomOpusApp = OpusAppWindowClass;

            _classAtom = Atom.Zero;

            _inGetWindow = 1;
            _hWndWindow = IntPtr.Zero;

            if (_wordApplication.Windows.Count == 0)
            {
                _classAtom = classAtomOpusApp;
                _wordApplication.Caption = _wordApplication.Caption;
                
                hWnd = _hWndWindow;
            }
            else
            {
                if (window == null)
	                window = _wordApplication.ActiveWindowLJ();


                if (window != null)
                {
                    //hWnd = GetVirticalPaneScrollbarHandle(window.Panes(1)) 'May throw Exception "This object model command is not available while in the current event."
                    hWnd = GetVerticalPaneScrollbarHandle(window.ActivePane);
                    if (IsWindow(hWnd) == false)
                    {
                        hWnd = FindWordWindowHandle(window);
                    }
                }
            }

            if (IsWindow(hWnd))
            {
                var hWndParent = hWnd;
                var classAtomToCompare = new Atom();
                while (!(classAtomOpusApp == classAtomToCompare))
                {
                    hWnd = GetParent(hWndParent);
                    if (!IsWindow(hWnd))
                    {
                        break;
                    }

                    classAtomToCompare = GetClassAtom(hWnd);
                    hWndParent = hWnd;
                }

                if (IsWindow(hWndParent))
                {
                    return hWndParent;
                }
            }

            return default;
        }


        //Finds the correct opus app Window handle from a Window object
        // This method uses the FindWindowEx api to loop through all the OpusApp windows then drilling down to the
        // _WwG window which we can call the IAccessibility AccessibleObjectFromWindow which will return a Window instance.
        // When the Window returned from AccessibleObjectFromWindow matches the Window passed in, we know we have the correct OpusApp window handle.
        private IntPtr FindWordWindowHandle(Window window)
        {
            //Find OpusApp
            var hWndOpusApp = FindWindowEx(IntPtr.Zero, IntPtr.Zero, OpusAppWindowClass, lpszWindow: null);
            while (hWndOpusApp != IntPtr.Zero)
            {
                //Find WwF
                var hWndWwF = FindWindowEx(hWndOpusApp, IntPtr.Zero, WwFWindowClass, lpszWindow: null);
                if (hWndWwF != IntPtr.Zero)
                {
                    var hWndWwB = FindWindowEx(hWndWwF, IntPtr.Zero, WwBWindowClass, lpszWindow: null);
                    while (hWndWwB != IntPtr.Zero)
                    {
                        //Find WwB

                        if (hWndWwB != IntPtr.Zero)
                        {
                            Window win = null;
                            var hWndWwG = FindWindowEx(hWndWwB, IntPtr.Zero, WwGWindowClass, lpszWindow: null);
                            if (TryGetWindowFromHandle(hWndWwG, ref win) && win == window)
                            {
                                return hWndWwG;
                            }
                        }

                        hWndWwB = FindWindowEx(hWndWwF, hWndWwB, WwBWindowClass, lpszWindow: null);
                    }
                }

                hWndOpusApp = FindWindowEx(IntPtr.Zero, hWndOpusApp, OpusAppWindowClass, lpszWindow: null);
            }

            return IntPtr.Zero;
        }


        //Public Function GetMDIClientWindowHandle() As System.IntPtr
        //  Return GetMDIClientWindowHandle(DirectCast(Nothing, Window))
        //End Function
        public IntPtr GetMdiClientWindowHandle(Window wordWindow)
        {
            return GetMdiClientWindowHandleInternal(wordWindow);
        }

        [SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic")]
        public IntPtr GetMdiClientWindowHandle(IntPtr hWndApplication)
        {
            return GetMdiClientWindowHandleInternal(hWndApplication);
        }

        private IntPtr GetMdiClientWindowHandleInternal(Window wordWindow)
        {
            var hWndApplication = GetApplicationHandleInternal(wordWindow);
            if (hWndApplication != IntPtr.Zero)
            {
                return GetMdiClientWindowHandle(hWndApplication);
            }

            return IntPtr.Zero;
        }

        private static IntPtr GetMdiClientWindowHandleInternal(IntPtr hWndApplication)
        {
            if (hWndApplication == IntPtr.Zero)
            {
                return IntPtr.Zero;
            }

            var hWnd = FindWindowEx(hWndApplication, IntPtr.Zero, WwFWindowClass, lpszWindow: null);
            while (!(hWnd == IntPtr.Zero))
            {
                if (IsWindowVisible(hWnd))
                {
                    if (GetParent(hWnd) == hWndApplication)
                    {
                        break;
                    }
                }

                hWnd = FindWindowEx(hWndApplication, hWnd, WwFWindowClass, lpszWindow: null);
            }

            return hWnd;
        }


        public IntPtr GetDocumentWindowHandle(Window window)
        {
            Check.NotNull(window, nameof(window));
            return GetDocumentWindowHandleInternal(window);
        }

        private IntPtr GetDocumentWindowHandleInternal(Window window)
        {
            var hWnd = GetVerticalPaneScrollbarHandle(window.Panes[Index: 1]);
            if (IsWindow(hWnd) == false)
            {
                hWnd = FindWordWindowHandle(window);
            }

            if (IsWindow(hWnd))
            {
                hWnd = GetParent(hWnd);
                if (IsWindow(hWnd))
                {
                    if (GetClassAtom(hWnd) == GetClassAtomFromName("_WwB"))
                    {
                        return hWnd;
                    }
                }
            }

            return IntPtr.Zero;
        }


        private static bool TryGetWindowFromHandle(IntPtr handle, ref Window window)
        {
            if (IsParentOfRealOpusApp(handle))
            {
                window = GetObjectFromWindow<Window>(handle);
            }

            try
            {
                if (window != null && window.Document.Kind == WdDocumentKind.wdDocumentEmail)
                {
                    window = null;
                }
            }
            catch (COMException)
            {
                // Ignore
            }

            return window != null;
        }


        private static bool IsParentOfRealOpusApp(IntPtr handle)
        {
            var wi = new WindowInfo(handle).Root;
            if (wi != null && wi.IsValid)
            {
                if (GetClassAtom(wi.Handle) == OpusAppWindowClass)
                {
                    return !string.IsNullOrEmpty(wi.Text);
                }
            }

            return false;
        }


        public static Window GetWindowFromHandle(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero)
            {
                return null;
            }

            return GetObjectFromWindow<Window>(hWnd);
        }


        private static T GetObjectFromWindow<T>(IntPtr hWnd) where T : class
        {
            object obj = null;

            hWnd = NormalizeHandleToWwG(hWnd);

            if (IsWindow(hWnd))
            {
                //We have a window handle to the _WwG window
                var iid = new Guid(a: 0x20400, b: 0, c: 0, d: 0xC0, e: 0, f: 0, g: 0, h: 0, i: 0, j: 0, k: 0x46);
                //IDispatch
                var hr = AccessibleObjectFromWindow(hWnd, OBJID_NATIVEOM, ref iid, ref obj);
                if (hr != 0)
                {
                    Marshal.GetExceptionForHR(hr);
                }

                return (T)obj;
            }

            return default;
        }


        private void SetHooks(bool hook)
        {
            if (_hooked == hook)
                return;

            if (hook)
            {
                WindowsHook.Add(WindowsHookType.WH_CALLWNDPROC, CallWndProc);
                _hooked = true;
                return;
            }

            WindowsHook.Remove(WindowsHookType.WH_CALLWNDPROC, CallWndProc);
            _hooked = false;
        }


        private IntPtr GetVerticalPaneScrollbarHandle(Pane pane)
        {
            if (decimal.TryParse(_wordApplication.Version, out var version))
            {
                return GetVerticalPaneScrollbarHandle(pane, decimal.ToInt32(version));
            }

            return IntPtr.Zero;
        }

        private IntPtr GetVerticalPaneScrollbarHandle(Pane pane, int officeVersion)
        {
            IntPtr hWnd;
            _hWndWindow = IntPtr.Zero;
            _inGetWindow = 1;

            try
            {
                _classAtom = GetClassAtomFromName(officeVersion < 12 ? "ScrollBar" : "NUIScrollBar");
                // ReSharper disable once UnusedVariable
                var value = pane.VerticalPercentScrolled;

                hWnd = _hWndWindow;
            }
            finally
            {
                _inGetWindow = 0;
                _classAtom = Atom.Zero;
                _hWndWindow = IntPtr.Zero;
            }

            return hWnd;
        }

        private static bool IsWwGDocumentWindow(IntPtr handle)
        {
            if (IsWindowClass(handle, WwGWindowClass))
            {
                var winInfo = new WindowInfo(handle);
                if ((winInfo.ExStyle & WindowExStyles.WS_EX_ACCEPTFILES) == WindowExStyles.WS_EX_ACCEPTFILES &&
                    (winInfo.Style & WindowStyles.WS_VISIBLE) == WindowStyles.WS_VISIBLE)
                {
                    return true;
                }
            }

            return false;
        }



        [DebuggerStepThrough]
        private void WmSize(ref CwpStruct cwpStruct)
        {
            //Debug.WriteLine(GetClassName(cwpStruct.Handle))
            var state = cwpStruct.WParam.ToInt32();
            switch (state)
            {
            case 0: //Normal:
                case 1://Minimized:
                case 2://Maximized:

                if (GetClassAtom(cwpStruct.Handle) == OpusAppWindowClass ||
                    GetClassAtom(cwpStruct.Handle) == WwBWindowClass)
                {
                    Window window = null;
                    if (TryGetWindowFromHandle(cwpStruct.Handle, ref window))
                    {
                        WindowStateChanged?.Invoke(this,
                                                   new WordWindowStateChangedEventArgs(window, cwpStruct.Handle,
                                                                                       state));
                    }
                }

                break;
            }
        }


        private void WmSysColorChange(ref CwpStruct cwpStruct)
        {
            if (IsWindowClass(cwpStruct.Handle, OpusAppWindowClass) || IsWindowClass(cwpStruct.Handle, MsoWorkPaneWindowClass))
            {
                var theme = _wordApplication.GetOfficeTheme();
                if (Theme != theme)
                {
                    _theme = theme;
                    WordThemeChanged?.Invoke(this, EventArgs.Empty);
                }
            }
        }

        private void WmActivate(ref CwpStruct cwpStruct)
        {
            if (_windowsInTaskBar == WordBoolean.False)
	            return;


            if (WindowActivateChanged == null)
	            return;

            Window window = null;
            if (TryGetWindowFromHandle(cwpStruct.Handle, ref window))
            {
	            var active = cwpStruct.WParam.ToInt32() != WA_INACTIVE;
                var e = new WordWindowActivateEventArgs(window, cwpStruct.Handle, false, active);
                WindowActivateChanged.Invoke(this, e);
            }
        }

        private void WmNcActivate(ref CwpStruct cwpStruct)
        {
            if (_windowsInTaskBar == WordBoolean.True)
	            return;

            if (WindowActivateChanged == null)
	            return;

            Window window = null;

            if (TryGetWindowFromHandle(cwpStruct.Handle, ref window))
            {

	            var active = cwpStruct.WParam == IntPtr.Zero;
	            var e = new WordWindowActivateEventArgs(window, cwpStruct.Handle, true, active);
	            WindowActivateChanged.Invoke(this, e);
            }
        }

		[SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
		private void WmParentNotify(ref CwpStruct cwpStruct)
		{

			var msg = LoWord(cwpStruct.WParam);
			if (msg != WindowMessages.WM_DESTROY)
				return;

			if (WordExtensions.Quitting)
				return;

			try
			{

				if (IsWindowClass(cwpStruct.Handle, WwBWindowClass) && IsWindowClass(cwpStruct.LParam, WwGWindowClass))
				{

					Window window = null;
					if (TryGetWindowFromHandle(cwpStruct.Handle, ref window))
					{
                        //BUG BA-332: Word Crashes in Review removing the Split bar while in Draft View.
                        //Also BA exists ReviewMode when switching to draft mode in ReviewMode 
                        //When a split window closes we are incorrectly a Window close event
                        //What we need to do is check if more than one _WwG window exists.
                        // if we do then don't raise the WindowClosed event yet.

                        var wi = new WindowInfo(cwpStruct.Handle); //_WwB
                        var countOfWwG = 0;
						wi.ProcessChildWindows(directDescendentsOnly:true, info =>
						{
							if (IsWindowClass(info.Handle, WwGWindowClass))
								countOfWwG++;
							return countOfWwG < 2; //continue if less than 2, end processing if 2 or more exists
                        });
						if (countOfWwG > 1)
							return; //More than one _WwF window still exists so don't close.


						WindowClosed?.Invoke(this, new WordWindowEventArgs(window, cwpStruct.Handle));
                    }
				}

			}
			catch (Exception ex)
			{
				Debug.WriteLine("Exception in WmDestroy");
			}

		}


		private void WmSetFocus(ref CwpStruct cwpStruct)
        {
            //debug.WriteLine((New WindowInfo(cwpStruct.Handle)).ToString)
            FocusChanged?.Invoke(this, new FocusChangedEventArgs(cwpStruct.WParam, cwpStruct.Handle));
        }


        private void WmKillFocus(ref CwpStruct cwpStruct)
        {
            FocusLost?.Invoke(this, new FocusChangedEventArgs(cwpStruct.WParam, cwpStruct.Handle));
        }

 
        [DebuggerNonUserCode]
        private void CallWndProc(ref HookMessage m)
        {
            var cwpStruct = CwpStruct.NewCwpStruct(ref m);
    
            if (_inGetWindow == 1)
            {
                if (_classAtom == GetClassAtom(cwpStruct.Handle))
                {
                    _inGetWindow = 2;
                    _hWndWindow = cwpStruct.Handle;
                }
            }
             
            m.CallNextHook();
 
            switch (cwpStruct.Msg)
            {
            case WindowMessages.WM_SETFOCUS:
                    WmSetFocus(ref cwpStruct);
                    break;
                case WindowMessages.WM_KILLFOCUS:
                    WmKillFocus(ref cwpStruct);
                    break;
                case WindowMessages.WM_SYSCOLORCHANGE:
                    WmSysColorChange(ref cwpStruct);
                    break;
                case WindowMessages.WM_SIZE:
                    WmSize(ref cwpStruct);
                    break;
                case WindowMessages.WM_ACTIVATE:
                    WmActivate(ref cwpStruct);
                    break;
                case WindowMessages.WM_NCACTIVATE:
                    WmNcActivate(ref cwpStruct);
                    break;
                case WindowMessages.WM_SETTEXT:
	                break;
                case WindowMessages.WM_PARENTNOTIFY:
                    WmParentNotify(ref cwpStruct);
                    break;

            }
        }


        private bool _disposedValue; // To detect redundant calls

 
        ~WordWindowHelper()
        {
#if (TRACK_DISPOSED)
		    LevitJames.Core.Diagnostics.DisposeTracker.Add(_disposedSource);
#endif
            Dispose();
        }

        public void Dispose()
        {
            if (!_disposedValue)
            {

                SetHooks(hook: false);
                AppDiagnostics.OptionChanged -= AppDiagnosticsChangedHandler;

                _wordApplication = null;
            }

            _disposedValue = true;
            GC.SuppressFinalize(this);
        }
    }
}