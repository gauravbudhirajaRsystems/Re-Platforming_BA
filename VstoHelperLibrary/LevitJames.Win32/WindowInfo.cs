//*************************************************
//* © 2020 Litera Corp. All Rights Reserved.
//**************************************************

using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace LevitJames.Win32
{

    [Flags]
    public enum SetBoundsFlags
    {
        None = 0,
        IgnoreLocation = 0x1,
        IgnoreSize = 0x2,
        Activate = 0x4,
        FrameChanged = 0x8
    }

    /// <summary>
    ///     A class for obtaining information about a window handle.
    /// </summary>
    [Serializable]
    public class WindowInfo : ICloneable
    {

        [DebuggerStepThrough]
        public WindowInfo(IntPtr handle)
        {
            Handle = handle;
        }

        [DebuggerStepThrough]
        public WindowInfo(IntPtr? handle)
        {
            if (handle.HasValue)
            {
                Handle = handle.Value;
            }
        }

        public int Dpi => NativeMethods.GetDpiForWindow(Handle);

        public void SendBeforeDpiChanged(int expectedDpi = 0) => SendDpiChanged(expectedDpi, before: true);
        public void SendAfterDpiChanged(int expectedDpi = 0) => SendDpiChanged(expectedDpi, before: false);

        public void SendBeforeAfterDpiChanged(int expectedDpi = 0)
        {
            SendDpiChanged(expectedDpi, before: true);
            SendDpiChanged(expectedDpi, before: false);
        }

        private void SendDpiChanged(int expectedDpi, bool before)
        {

            var window = new WindowInfo(Handle);
            if (expectedDpi > 0 && window.Dpi == expectedDpi)
                return;

            var msg = before ? WindowMessages.WM_DPICHANGED_BEFOREPARENT : WindowMessages.WM_DPICHANGED_AFTERPARENT;
            window.SendMessage(msg, IntPtr.Zero, IntPtr.Zero);
            window.ProcessChildWindows((w) =>
            {

                w.SendMessage(msg, IntPtr.Zero, IntPtr.Zero);
                return true;
            });
        }
        public bool IsActiveWindow => IsValid && Handle == NativeMethods.GetActiveWindow();

        public IntPtr GetWindowLong(int nIndex) => NativeMethods.GetWindowLong(HandleRef, nIndex);

        public int GetWindowLongInt32(int nIndex) => NativeMethods.GetWindowLongInt32(HandleRef, nIndex);
        public IntPtr SetWindowLong(int nIndex, int dwNewLong) => NativeMethods.SetWindowLong(HandleRef, nIndex, dwNewLong);
        public IntPtr SetWindowLong(int nIndex, IntPtr dwNewLong) => NativeMethods.SetWindowLong(HandleRef, nIndex, dwNewLong);

        /// <summary>
        ///     Returns a Boolean flag specifying of the handle for this instance is valid.
        /// </summary>
        public bool IsValid => NativeMethods.IsWindow(HandleRef);


        [SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly", MessageId = "InCurrent")]
        public bool InCurrentProcess
        {
            get
            {
                if (IsValid)
                {
                    return NativeMethods.IsWindowInCurrentProcess(HandleRef);
                }

                return false;
            }
        }

        public IntPtr GetDeviceContext(bool client = true)
        {

            if (IsValid)
            {
                if (client)
                    return NativeMethods.GetDC(HandleRef);

                return NativeMethods.GetWindowDC(HandleRef);
            }

            return IntPtr.Zero;

        }

        public void ReleaseDeviceContext(IntPtr hDc)
        {

            if (IsValid && hDc != IntPtr.Zero)
                _ = NativeMethods.ReleaseDC(HandleRef, hDc);

        }
        /// <summary>
        ///     Returns the a WindowInfo instance representing the Parent window
        /// </summary>
        /// <returns>Returns Null if there is no valid parent.</returns>
        public WindowInfo Parent => new WindowInfo(NativeMethods.GetParent(Handle));



        /// <summary>
        ///     Checks if the Window is a child window or a floating window.
        /// </summary>
        /// <returns>True, if the window style contains the WS_CHILD; otherwise false.</returns>
        public bool IsChildWindow
            => ((WindowStyles)Style & WindowStyles.WS_CHILD) == WindowStyles.WS_CHILD;


        /// <summary>
        ///     Returns the a WindowInfo instance by walking the chain of parent windows.
        /// </summary>
        /// <returns>Returns Null if there is no valid root window.</returns>
        public WindowInfo Root
        {
            get
            {
                var window = new WindowInfo(NativeMethods.GetAncestor(Handle, NativeMethods.GAFlags.Root));
                if (window.IsValid)
                {
                    return window;
                }

                return null;
            }
        }


        /// <summary>
        ///     Returns the a WindowInfo instance by walking the chain of parent and owner windows.
        /// </summary>
        /// <returns>Returns Null if there is no valid root window.</returns>
        public WindowInfo RootOwner
        {
            get
            {
                var window = new WindowInfo(NativeMethods.GetAncestor(Handle, NativeMethods.GAFlags.RootOwner));
                if (window.IsValid)
                {
                    return window;
                }

                return null;
            }
        }


        /// <summary>
        ///     Returns the a WindowInfo instance by walking the chain of parent and owner windows.
        /// </summary>
        /// <returns>Returns Null if there is no valid root window.</returns>
        public WindowInfo Owner
        {
            get
            {
                var window =
                    new WindowInfo(NativeMethods.GetWindowLong(HandleRef, NativeMethods.GWL_HWNDPARENT));
                return window.IsValid ? window : null;
            }
        }

        public WindowInfo ChildOwnerWindow //=> new WindowInfo(NativeMethods.GetWindow(Handle, NativeMethods.GetWindowCommand.GW_OWNER));
        {
            get
            {
                var window = new WindowInfo(NativeMethods.GetTopWindow(Handle));
                return window.IsValid ? window : null;
            }

        }

        public WindowInfo PreviousWindow
            => new WindowInfo(NativeMethods.GetWindow(Handle, NativeMethods.GetWindowCommand.GW_HWNDPREV));

        public WindowInfo AboveWindow => PreviousWindow;


        public WindowInfo NextWindow
            => new WindowInfo(NativeMethods.GetWindow(Handle, NativeMethods.GetWindowCommand.GW_HWNDNEXT));

        public WindowInfo BelowWindow => NextWindow;


        public WindowInfo FirstWindow
            => new WindowInfo(NativeMethods.GetWindow(Handle, NativeMethods.GetWindowCommand.GW_HWNDFIRST));


        public WindowInfo LastWindow
            => new WindowInfo(NativeMethods.GetWindow(Handle, NativeMethods.GetWindowCommand.GW_HWNDLAST));


        public WindowInfo ChildWindow
            => new WindowInfo(NativeMethods.GetWindow(Handle, NativeMethods.GetWindowCommand.GW_CHILD));


        /// <summary>
        ///     Returns the class name of the window
        /// </summary>
        public string ClassName => NativeMethods.GetClassName(Handle);


        public int ClassAtom
        {
            get
            {
                if (!IsValid)
                {
                    return 0;
                }

                return Convert.ToInt32(NativeMethods.GetClassAtom(HandleRef));
            }
        }


        /// <summary>
        ///     Returns the window text (caption) of the window
        /// </summary>
        public string Text => NativeMethods.GetWindowText(Handle);

        /// <summary>
        /// Moves a window like it is moved when the titlebar is moved
        /// </summary>
        public void SimulateTitleBarMove()
        {
            if (!IsValid)
                return;
            NativeMethods.ReleaseCapture();
            NativeMethods.SendMessage(HandleRef, WindowMessages.WM_NCLBUTTONDOWN,
                (IntPtr)NativeMethods.HitTestResult.HTCAPTION, IntPtr.Zero);
        }

        /// <summary>
        ///     Returns if the window is visible
        /// </summary>
        public bool Visible => NativeMethods.IsWindowVisible(HandleRef);


        /// <summary>
        ///     Returns if the window is enabled
        /// </summary>
        public bool Enabled
        {
            get => NativeMethods.IsWindowEnabled(HandleRef);
            set
            {
                if (IsValid)
                {
                    NativeMethods.EnableWindow(Handle, value);
                }
            }
        }

        public WindowPlacement GetWindowPlacement()
        {
            if (!IsValid)
                throw new Exception("Invalid Window handle");

            var wp = new WindowPlacement();
            NativeMethods.GetWindowPlacement(Handle, wp);
            return wp;


        }
        public void SetWindowPlacement(WindowPlacement wp)
        {
            if (wp == null)
                throw new ArgumentNullException(nameof(wp));

            if (!IsValid)
                throw new Exception("Invalid Window handle");

            NativeMethods.SetWindowPlacement(Handle, wp);
        }

        /// <summary>
        ///     Returns the window style flags for the window.
        /// </summary>
        public WindowStyles Style
        {
            get
            {
                if (IsValid)
                {
                    return (WindowStyles)NativeMethods.GetWindowLongInt32(HandleRef, NativeMethods.GWL_STYLE);
                }

                return 0;
            }
        }


        /// <summary>
        ///     Returns the extended window style flags for the window.
        /// </summary>
        public WindowExStyles ExStyle => (WindowExStyles)NativeMethods.GetWindowLongInt32(HandleRef, NativeMethods.GWL_EXSTYLE);

        public void UpdateStyle(WindowStyles style, bool updateWindow = true)
            => UpdateStyle((int)style, false, updateWindow);

        public void UpdateExStyle(WindowExStyles style, bool updateWindow = true)
            => UpdateStyle((int)style, true, updateWindow);

        private void UpdateStyle(int style, bool exStyle, bool updateWindow)
        {
            if (!IsValid)
                return;
            var index = exStyle ? NativeMethods.GWL_EXSTYLE : NativeMethods.GWL_STYLE;

            NativeMethods.SetWindowLong(HandleRef, index, style);

            if (updateWindow)
                NativeMethods.SetWindowPos(Handle, IntPtr.Zero, 0, 0, 0, 0,
                    SetWindowPosFlags.SWP_NOSIZE_NOMOVE_NOACTIVATE | SetWindowPosFlags.SWP_FRAMECHANGED);
        }


        /// <summary>
        ///     Returns the rectangle for the window in screen coordinates
        /// </summary>
        public RectangleI Bounds => NativeMethods.GetWindowRect(HandleRef);


        /// <summary>
        ///     Returns the rectangle for the window in client coordinates
        /// </summary>
        public RectangleI ClientBounds => NativeMethods.GetClientRect(HandleRef);

        public RectangleI LocalBounds()
        {
            var parent = Parent;
            if (!IsChildWindow)
                return Bounds;

            var localBounds = ClientBounds;
            localBounds = NativeMethods.MapWindowPoints(Handle, Parent.Handle, localBounds);
            return localBounds;
        }

        /// <summary>
        ///     Returns the size of the window
        /// </summary>
        public SizeI Size => NativeMethods.GetWindowRect(HandleRef).Size;


        /// <summary>
        ///     Returns the location of the window
        /// </summary>
        public PointI Location => NativeMethods.GetWindowRect(HandleRef).Location;


        public HandleRef HandleRef => new HandleRef(this, Handle);

        object ICloneable.Clone()
        {
            return Clone();
        }


        /// <summary>
        ///     Returns a win32 window handle
        /// </summary>
        public IntPtr Handle { get; }


        public static WindowInfo FromObject(object objWindow)
        {
            if (objWindow == null)
                throw new ArgumentNullException(nameof(objWindow));

            var handle = default(IntPtr);
            if (objWindow is IntPtr win32)
            {
                handle = win32;
            }
            else if (objWindow is IOleWindow oleWindow)
            {
                handle = oleWindow.GetWindow;
            }

            return new WindowInfo(handle);
        }


        public static WindowInfo FromDesktop() => new WindowInfo(NativeMethods.GetDesktopWindow());

        public static WindowInfo FromFocus() => new WindowInfo(NativeMethods.GetFocus());

        public static WindowInfo FromForegroundWindow() => new WindowInfo(NativeMethods.GetForegroundWindow());

        public static WindowInfo FromActiveWindow() => new WindowInfo(NativeMethods.GetActiveWindow());

        public static WindowInfo FromPoint(int screenLocationX, int screenLocationY) => new WindowInfo(NativeMethods.WindowFromPoint(screenLocationX, screenLocationY));

        public static WindowInfo FromPoint(PointI screenLocation) => new WindowInfo(NativeMethods.WindowFromPoint(screenLocation.X, screenLocation.Y));

        public static WindowInfo FromMousePosition()
        {
            var pt = NativeMethods.GetCursorPos();
            return FromPoint(pt.Y, pt.Y);
        }

        public static WindowInfo FromMouseCapture() => new WindowInfo(NativeMethods.GetCapture());


        public static WindowInfo FromCaretWindow()
        {
            var guiThreadInfo = NativeMethods.GUITHREADINFO.NewGUITHREADINFO();
            var threadId = NativeMethods.GetCurrentWin32ThreadId();
            NativeMethods.GetGUIThreadInfo(threadId, ref guiThreadInfo);
            return new WindowInfo(guiThreadInfo.hwndCaret);
        }


        public static WindowInfo FindWindow(string className, string text)
        {
            return new WindowInfo(NativeMethods.FindWindowEx(IntPtr.Zero, IntPtr.Zero, className, text));
        }

        public static WindowInfo FindWindow(IntPtr parent, IntPtr windowAfter, string className, string text)
        {

            return new WindowInfo(NativeMethods.FindWindowEx(parent, windowAfter, className, text));
        }


        public int ThreadId() => NativeMethods.GetWindowThreadProcessId(HandleRef, out var _);


        public bool InCurrentThread() => !IsValid ? false : ThreadId() == NativeMethods.GetCurrentThreadId();

        /// <summary>
        ///     Gets the process id the window belongs too.
        /// </summary>
        public int ProcessId()
        {
            int retVal;
            NativeMethods.GetWindowThreadProcessId(HandleRef, out retVal);
            return retVal;
        }


        public bool ContainsFocus()
        {
            var window = FromFocus();
            while (window.IsValid)
            {
                if (window == this)
                {
                    return true;
                }

                window = window.Parent;
            }

            return false;
        }


        public WindowInfo GetChildFromDialogId(int id)
        {
            var wi = new WindowInfo(NativeMethods.GetDlgItem(HandleRef, id));
            return wi;
        }


        public void Invalidate()
        {
            Invalidate(new RectangleI(), children: false, erase: false, frame: false);
        }

        public void Invalidate(RectangleI updateRectangle, bool children, bool erase, bool frame)
        {
            if (IsValid)
            {
                var flags = NativeMethods.RedrawWindowFlags.RDW_INVALIDATE |
                            (children
                                 ? NativeMethods.RedrawWindowFlags.RDW_ALLCHILDREN
                                 : NativeMethods.RedrawWindowFlags.RDW_NONE) |
                            (erase
                                 ? NativeMethods.RedrawWindowFlags.RDW_ERASE
                                 : NativeMethods.RedrawWindowFlags.RDW_NONE) |
                            (frame
                                 ? NativeMethods.RedrawWindowFlags.RDW_FRAME
                                 : NativeMethods.RedrawWindowFlags.RDW_NONE);

                if (updateRectangle.IsEmpty)
                {
                    NativeMethods.RedrawWindow(Handle, IntPtr.Zero, IntPtr.Zero, flags);
                }
                else
                {
                    NativeMethods.RedrawWindow(Handle, ref updateRectangle, IntPtr.Zero, flags);
                }
            }
        }


        /// <summary>
        ///     returns if two WindowInfo objects point to the same Win32 window
        /// </summary>
        /// <param name="obj"></param>
        public bool Equals(WindowInfo obj)
        {
            if (obj == null)
            {
                return false;
            }

            return obj.Handle == Handle;
        }



        public void EnumWindowThreadWindows(Predicate<WindowInfo> predicate)
        {
            NativeMethods.EnumThreadWindows(ThreadId(),
                                            (hWnd, lParam) => predicate(new WindowInfo(hWnd)), default(HandleRef));
        }

        public static void EnumThreadWindows(Predicate<WindowInfo> predicate, int threadId = 0)
        {
            NativeMethods.EnumThreadWindows(threadId != 0 ? threadId : NativeMethods.GetCurrentWin32ThreadId(),
                                            (hWnd, lParam) => predicate(new WindowInfo(hWnd)), default(HandleRef));
        }


        public void ProcessChildWindows(Predicate<WindowInfo> predicate)
        {
            ProcessChildWindows(directDescendentsOnly: false, predicate: predicate);
        }

        public void ProcessChildWindows(bool directDescendentsOnly, Predicate<WindowInfo> predicate)
        {
            NativeMethods.EnumChildWindows(HandleRef, (hWnd, lParam) =>
            {
                var child = new WindowInfo(hWnd);
                var add = true;
                if (directDescendentsOnly)
                {
                    add = child.Parent.Handle == Handle;
                }

                if (add)
                {
                    return predicate(child);
                }

                return true;
            }, new IntPtr(value: 0));
        }

        public void Update()
        {
            if (!IsValid)
                return;
            NativeMethods.UpdateWindow(Handle);
        }

        public void Refresh()
        {
            if (!IsValid)
                return;

            NativeMethods.RedrawWindow(Handle, IntPtr.Zero, IntPtr.Zero,
                                       NativeMethods.RedrawWindowFlags.RDW_ERASE |
                                       NativeMethods.RedrawWindowFlags.RDW_FRAME |
                                       NativeMethods.RedrawWindowFlags.RDW_ALLCHILDREN |
                                       NativeMethods.RedrawWindowFlags.RDW_UPDATENOW);
        }


        /// <summary>
        ///     Suspends painting the Form to maintain performance when many controls need to be updated.
        /// </summary>
        /// <remarks>
        ///     The number of calls to BeginUpdate must be matched to the same number of calls to EndUpdate to resume
        ///     painting.
        /// </remarks>
        public void BeginUpdate()
        {
            if (!IsValid)
            {
                return;
            }

            if (!Visible)
                return; // if the window is not visible then it will become visible in EndUpdate.

            NativeMethods.SendMessage(HandleRef, WindowMessages.WM_SETREDRAW, IntPtr.Zero,
                                      IntPtr.Zero);

        }


        /// <summary>
        ///     Resumes painting the Form after painting is suspended by the BeginUpdate method.
        /// </summary>
        public void EndUpdate()
        {
            EndUpdate(invalidate: true);
        }

        /// <summary>
        ///     Resumes painting the Form after painting is suspended by the BeginUpdate method.
        /// </summary>
        /// <param name="invalidate">Set to true to invalidate the window for redrawing.</param>

        public void EndUpdate(bool invalidate)
        {

            if (!IsValid)
                return;

            if (!Visible)
                return; // if the window is not visible then it will become visible.

            NativeMethods.SendMessage(HandleRef, WindowMessages.WM_SETREDRAW, new IntPtr(-1),
                                      IntPtr.Zero);
            if (invalidate)
            {
                NativeMethods.RedrawWindow(Handle, IntPtr.Zero, IntPtr.Zero,
                                           NativeMethods.RedrawWindowFlags.RDW_ERASE |
                                           NativeMethods.RedrawWindowFlags.RDW_FRAME |
                                           NativeMethods.RedrawWindowFlags.RDW_ALLCHILDREN |
                                           NativeMethods.RedrawWindowFlags.RDW_INVALIDATE);
                //UnManagedMethods.InvalidateRect(Me.Handle, IntPtr.Zero, 0)
            }
        }

        public void SetFocus()
        {
            if (IsValid)
            {
                NativeMethods.SetFocus(HandleRef);
            }
        }


        [SuppressMessage("Microsoft.Usage", "CA1806:DoNotIgnoreMethodResults",
            MessageId = "LevitJames.Libraries.NativeMethods.SetActiveWindow(System.Runtime.InteropServices.HandleRef)")]
        public void Activate()
        {
            if (IsValid)
            {
                NativeMethods.SetActiveWindow(HandleRef);
            }
        }


        public void SetBounds(int x, int y, int width, int height, SetBoundsFlags flags)
        {
            if (IsValid)
            {
                var swpf = SetWindowPosFlags.SWP_NOACTIVATE;
                swpf |= (flags & SetBoundsFlags.Activate) == SetBoundsFlags.Activate
                            ? SetWindowPosFlags.None
                            : SetWindowPosFlags.SWP_NOACTIVATE;
                swpf |= (flags & SetBoundsFlags.IgnoreLocation) == SetBoundsFlags.IgnoreLocation
                            ? SetWindowPosFlags.SWP_NOMOVE
                            : SetWindowPosFlags.None;
                swpf |= (flags & SetBoundsFlags.IgnoreSize) == SetBoundsFlags.IgnoreSize
                            ? SetWindowPosFlags.SWP_NOSIZE
                            : SetWindowPosFlags.None;
                swpf |= (flags & SetBoundsFlags.FrameChanged) == SetBoundsFlags.FrameChanged
                            ? SetWindowPosFlags.SWP_FRAMECHANGED
                            : SetWindowPosFlags.None;
                if (!InCurrentThread())
                    swpf |= SetWindowPosFlags.SWP_ASYNCWINDOWPOS;

                NativeMethods.SetWindowPos(Handle, IntPtr.Zero, x, y, width, height, swpf);
            }
        }

        public bool IsMinimized => (Style & WindowStyles.WS_MINIMIZE) == WindowStyles.WS_MINIMIZE;
        public bool IsMaximized => (Style & WindowStyles.WS_MAXIMIZE) == WindowStyles.WS_MAXIMIZE;

        public void Restore()
        {
            if (IsValid)
            {
                NativeMethods.ShowWindow(Handle, NativeMethods.ShowWindowFlags.SW_RESTORE);
            }
        }


        public void Maximize()
        {
            if (IsValid)
            {
                NativeMethods.ShowWindow(Handle, NativeMethods.ShowWindowFlags.SW_MAXIMIZE);
            }
        }


        public void Minimize()
        {
            if (IsValid)
            {
                NativeMethods.ShowWindow(Handle, NativeMethods.ShowWindowFlags.SW_MINIMIZE);
            }
        }


        public void Close()
        {
            //Root?.SendMessage(WindowMessages.WM_SYSCOMMAND, New IntPtr(SC_CLOSE), IntPtr.Zero)
            Root?.SendMessage(WindowMessages.WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        }


        public IntPtr PostMessage(int message, IntPtr wParam, IntPtr lParam)
        {
            if (IsValid)
            {
                return NativeMethods.PostMessage(HandleRef, message, wParam, lParam);
            }

            return default(IntPtr);
        }

        public void PostMessage(ref WindowMessage message)
        {
            if (IsValid)
            {
                message.Result = NativeMethods.PostMessage(HandleRef, message.Msg, message.WParam, message.LParam);
            }
        }

        public IntPtr SendMessage(int message, IntPtr wParam, IntPtr lParam)
        {
            if (IsValid)
            {
                return NativeMethods.SendMessage(HandleRef, message, wParam, lParam);
            }

            return default(IntPtr);
        }

        public void SendMessage(ref WindowMessage message)
        {
            if (IsValid)
            {
                message.Result = NativeMethods.SendMessage(HandleRef, message.Msg, message.WParam, message.LParam);
            }
        }


        public void SetToForeground()
        {
            if (IsValid)
            {
                NativeMethods.SetForegroundWindow(Handle);
            }
        }

        /// <summary>
        ///     Overloaded version of Form.SendToBack
        /// </summary>
        /// <param name="inFrontOfWindow">If Zero the window to brought to the back of all other windows; otherwise its brought to the back of the supplied window</param>
        /// <remarks>
        ///     The window is not activated.
        /// </remarks>
        public void SendToBack(bool activate = false, IntPtr inFrontOfWindow = default(IntPtr))
        {
            if (IsValid)
            {

                SetWindowPosFlags flags = SetWindowPosFlags.SWP_NOSIZE_NOMOVE_NOACTIVATE;
                if (inFrontOfWindow == default(IntPtr))
                    inFrontOfWindow = new IntPtr(NativeMethods.HWND_BOTTOM);
                else
                {
                    var thread = ThreadId();
                    var thread2 = NativeMethods.GetWindowThreadProcessId(new HandleRef(this, inFrontOfWindow), out var _);
                    if (thread2 != thread)
                    {
                        flags |= SetWindowPosFlags.SWP_ASYNCWINDOWPOS;
                    }
                }

                NativeMethods.SetWindowPos(Handle, inFrontOfWindow, x: 0, y: 0, cx: 0, cy: 0,
                                           wFlags: flags);
            }
        }

        /// <summary>
        /// Brings a form to the front of the ZOrder.
        /// </summary>
        /// <param name="activate">true to activate the window; false otherwise</param>
        /// <param name="inFrontOfWindow">If Zero the window to brought to the top of all other windows; otherwise its brought to the front of the supplied window</param>
        /// <remarks>
        ///     The window is not activated.
        /// </remarks>
        public void BringToFront(bool activate = false, IntPtr inFrontOfWindow = default(IntPtr))
        {
            if (IsValid)
            {
                var flags = activate ? SetWindowPosFlags.None : SetWindowPosFlags.SWP_NOACTIVATE;

                if (inFrontOfWindow == default(IntPtr))
                    inFrontOfWindow = new IntPtr(NativeMethods.HWND_TOP);
                else
                {
                    var thread = ThreadId();
                    var thread2 = NativeMethods.GetWindowThreadProcessId(HandleRef, out var _);
                    if (thread2 != thread)
                    {
                        flags |= SetWindowPosFlags.SWP_ASYNCWINDOWPOS;
                    }
                }


                NativeMethods.SetWindowPos(Handle, inFrontOfWindow, x: 0, y: 0, cx: 0, cy: 0,
                                           wFlags:
                                           SetWindowPosFlags.SWP_NOSIZE_NOMOVE | flags);
            }
        }

        public PointI PointToScreen(PointI clientPoint) => NativeMethods.MapWindowPoints(this.Handle, IntPtr.Zero, clientPoint);
        public RectangleI RectangleToScreen(RectangleI clientRectangle) => NativeMethods.MapWindowPoints(this.Handle, IntPtr.Zero, clientRectangle);

        public RectangleI RectangleFromScreen(RectangleI clientRectangle) => NativeMethods.MapWindowPoints(IntPtr.Zero, this.Handle, clientRectangle);
        public PointI PointFromScreen(PointI clientPoint) => NativeMethods.MapWindowPoints(IntPtr.Zero, this.Handle, clientPoint);

        /// <summary>
        /// Gets the MONITORINFOEX struct for the window, If defaultToPrimary is false then the default (if the window is off the screen) is the nearest
        /// </summary>
        /// <param name="defaultToPrimary"></param>
        /// <returns></returns>
        public MONITORINFOEX MonitorInfoForWindow(bool defaultToPrimary = true)
        {
            var hMon = NativeMethods.MonitorFromWindow(HandleRef,
                defaultToPrimary ? NativeMethods.MONITOR_DEFAULTTOPRIMARY : NativeMethods.MONITOR_DEFAULTTONEAREST);

            if (hMon == IntPtr.Zero)
                return null;

            var monInfo = new MONITORINFOEX();
            var r = NativeMethods.GetMonitorInfo(new HandleRef(this, hMon), monInfo);
            if (r == false)
                return null;
            return monInfo;
        }

        public int? Property(string name)
        {
            if (IsValid == false)
            {
                return null;
            }

            var propValue = NativeMethods.GetProp(HandleRef, name);
            if (Marshal.GetLastWin32Error() == 0)
            {
                return propValue;
            }

            return null;
        }
        //public void SetProperty(string name, string value)
        //{
        //    var hGlobalString = Marshal.StringToCoTaskMemAuto(value);

        //    NativeMethods.SetProp(this.HandleRef,name, hGlobalString);
        //    Marshal.FreeHGlobal(hGlobalString);
        //}
        //public string GetProperty(string name)
        //{
        //    var hGlobalString = NativeMethods.GetProp(this.HandleRef,name);
        //    var value = Marshal.PtrToStringAuto(hGlobalString);

        //    return value;
        //}

        public void Property(string name, int? value)
        {
            if (IsValid == false)
            {
                return;
            }

            if (string.IsNullOrEmpty(name))
                throw new ArgumentNullException(nameof(name));

            if (value.HasValue)
            {
                NativeMethods.SetProp(HandleRef, name, value.Value);
            }
            else
            {
                NativeMethods.RemoveProp(HandleRef, name);
            }
        }


        public IntPtr ToBitmap()
        {
            if (IsValid == false)
            {
                return IntPtr.Zero;
            }

            var bounds = Bounds;

            var hDcWindow = GetDeviceContext();
            var hDc = NativeMethods.CreateCompatibleDC(hDcWindow);
            var hBmp = NativeMethods.CreateCompatibleBitmap(hDc, bounds.Width, bounds.Height);
            var hBmpOld = NativeMethods.SelectObject(hDc, hBmp);
            NativeMethods.PrintWindow(HandleRef, hDc, nFlags: 0);
            NativeMethods.SelectObject(hDc, hBmpOld);
            //NativeMethods.DeleteObject(hBmp);
            NativeMethods.DeleteObject(hDc);
            ReleaseDeviceContext(hDcWindow);

            return hBmp;
        }


        //Friend Function Placement() As WINDOWPLACEMENT
        //	Dim wp As WINDOWPLACEMENT
        //	wp.length = Marshal.SizeOf(GetType(WINDOWPLACEMENT))
        //	GetWindowPlacement(Me.Handle, wp)
        //	Return wp
        //End Function


        public object GetObject(Guid guid)
        {
            if (!IsValid)
                return null;
            object obj = null;
            var hr = NativeMethods.AccessibleObjectFromWindow(Handle, NativeMethods.OBJID_NATIVEOM, ref guid, ref obj);
            if (!NativeMethods.Succeeded(hr))
                throw Marshal.GetExceptionForHR(hr);

            return obj;
        }


        public void FlushMessages()
        {
            if (!IsValid)
                return;

            var msg = new WindowMessage();
            var hr = HandleRef;
            while (NativeMethods.PeekMessage(ref msg, hr, 0, 0, remove: 1))
            {
                NativeMethods.TranslateMessage(ref msg);
                NativeMethods.DispatchMessage(ref msg);
            }
        }

        //Operator overloads


        public static bool operator ==(WindowInfo value1, WindowInfo value2)
        {
            var isNull1 = Equals(value1, null);
            var isNull2 = Equals(value2, null);
            if (isNull1 && isNull2)
                return true;

            if (isNull1 || isNull2)
                return false;


            return value1.Handle == value2.Handle;
        }

        public static bool operator ==(WindowInfo value1, IntPtr value2)
        {
            if (value1 == null)
                return false;

            return value1.Handle == value2;
        }

        public static bool operator ==(IntPtr value1, WindowInfo value2)
        {
            if (value2 == null)
                return false;

            return value1 == value2.Handle;
        }


        public static bool operator !=(WindowInfo value1, WindowInfo value2)
        {
            return !(value1 == value2);
        }

        public static bool operator !=(WindowInfo value1, IntPtr value2)
        {
            return !(value1 == value2);
        }


        public static bool operator !=(IntPtr value1, WindowInfo value2)
        {
            return !(value1 == value2);
        }


        public override bool Equals(object obj)
        {
            if (obj is IntPtr win32)
            {
                return this.Handle == win32;
            }

            if (obj is WindowInfo wi)
            {
                return this.Handle == wi.Handle;
            }

            return false;
        }


        public override int GetHashCode()
        {
            return NativeMethods.IntPtrToInt32(Handle);
        }


        public WindowInfo Clone()
        {
            return new WindowInfo(Handle); //Don't clone the window lock
        }


        public override string ToString()
        {
            if (Handle == IntPtr.Zero)
                return "None";

            return "Handle=" + Handle.ToString("x") + ", ClassName=" + ClassName + ", Text=" + Text + ", Visible=" +
                   Visible + ", Enabled=" + Enabled + ",Bounds=" + Bounds;
        }

        public void Capture()
        {
            NativeMethods.SetCapture(Handle);
        }
        public void ReleaseCapture()
        {
            NativeMethods.ReleaseCapture();
        }

    }

}