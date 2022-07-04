//*************************************************
//* © 2021 Litera Corp. All Rights Reserved.
//**************************************************

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using LevitJames.Core;
using LevitJames.Core.Diagnostics;
using LevitJames.MSOffice.Addin;
using LevitJames.Win32;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using static LevitJames.Libraries.NativeMethods;
using Application = Microsoft.Office.Interop.Word.Application;
using Font = System.Drawing.Font;
using Version = System.Version;

namespace LevitJames.MSOffice.MSWord
{
    /// <summary>
    ///     A singleton class containing extensions for Microsoft
    /// </summary>

    [SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling")]
    [DesignerCategory(@"Code")]
    public sealed class WordExtensions : IDisposable
    {

        /// <summary>
        ///     Raised when the Microsoft Word Selection changes.
        /// </summary>
        public static event EventHandler<WordSelectionEventArgs> SelectionChanged;

        /// <summary>
        ///     Raised when the Word 2007 (and beyond), application theme has changed.
        /// </summary>

        public static event EventHandler<EventArgs> WordThemeChanged;

        /// <summary>
        ///     Raised when a Word Document window has been closed.
        /// </summary>

        public static event EventHandler<WordWindowEventArgs> WordWindowClosed;

        /// <summary>
        ///     Raised when a Word Document window has been closed.
        /// </summary>

        public static event EventHandler<WordDocumentEventArgs> WordDocumentClosed;

        /// <summary>
        ///     Raised when a Word Window is immediately activated.
        /// </summary>
        /// <remarks>Word's Window Activate only activates after the mouse has been released.</remarks>
        public static event EventHandler<WordWindowActivateEventArgs> WordWindowActivateChanged;
 
        /// <summary>
        ///     Raised when the Word Application is about to close.
        /// </summary>

        public static event EventHandler<EventArgs> WordApplicationClosing;

        /// <summary>
        ///     Raised when the Word Application is about to close.
        /// </summary>

        public static event EventHandler<EventArgs> AddinDisconnected;

        /// <summary>
        ///     Raised when the Word Application is about to close.
        /// </summary>

        public static event EventHandler<OfficeShortcutKeyPressedEventArgs> ShortcutKeyPressed;
 

        public static event EventHandler WordDocumentChange;


        //private static DialogManager _dialogManager;
        private static IWordSelectionChange _wordSelectionChange;
        private static Font _font;
        private static string _addinPath;
        private static OfficeShortcutKeyCollection _shortcutKeys;
        private static bool _wordEventsAttached;

        private static LockCounter _focusChangeActionLock;
        private static WindowInfo _priorFocus;

        private static LockCounter _wordScreenUpdateLock;

        private WordUndoRecord _wordUndoRecord;
        private SuppressWordEvents _suppressWordAppEvents;

#if (TRACK_DISPOSED)
        private readonly string _disposedSource;
#endif

        // private constructor
        private WordExtensions()
        {
#if (TRACK_DISPOSED)
            _disposedSource = LevitJames.Core.Diagnostics.DisposeTracker.GetSource();
#endif
        }


        public static WordExtensions NewWordExtensions(Application wordApplication)
        {
            if (Singleton != null && WordApplication == wordApplication)
                return Singleton;

            Check.NotNull(wordApplication, nameof(wordApplication));

            Singleton = new WordExtensions();
            WordApplication = wordApplication;

            WordWindowHelper = WordWindowHelper.NewWordWindowHelper(WordApplication);

            AttachEventHandlers();

            return Singleton;
        }


        /// <summary>
        ///     Returns the singleton instance of this class.
        /// </summary>

        /// <returns>A WordExtensions instance.</returns>

        public static WordExtensions Singleton { get; private set; }


        /// <summary>
        ///     Returns a Microsoft Word Application object
        /// </summary>

        /// <returns>A Application instance</returns>
        public static Application WordApplication { get; private set; }

        public static string GetChannelText(bool includeGuid = true)
        {
            var channels = new Dictionary<string, string>
            {
                {"492350f6-3a01-4f97-b9c0-c7c6ddf67d60", "Monthly Channel"},
                {"7ffbc6bf-bc32-4f92-8982-f9dd17fd3114", "Semi-Annual Channel"},
                {"64256afe-f5d9-4f86-8936-8840a6a4f5be", "Monthly Channel (Targeted)"},
                {"b8f9b850-328d-4355-9145-c59439a0c4cf", "Semi-Annual Channel (Targeted)"},
                {"5440fd1f-7ecb-4221-8110-145efaa6372f", "Insider"}
            };

            const string channelSubKey = @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration";
            var valueName = "CDNBaseUrl";

            var valueText = string.Empty;
            using (var rootKey = IsX64()
                ? Registry.LocalMachine
                : RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
            {
                var value = rootKey.GetKeyValue(channelSubKey, valueName);
                if (value == null)
                {
                    valueName = "UpdateChannel";
                    value = rootKey.GetKeyValue(channelSubKey, valueName);
                    if (value != null)
                    {
                        valueText = value.ToString();
                        var lastSlash = valueText.LastIndexOf(@"/", StringComparison.InvariantCulture) + 1;
                        if (lastSlash > 0)
                            valueText = valueText.Substring(lastSlash);
                    }
                }
                else
                {
                    valueText = value.ToString();
                }
            }

            if (string.IsNullOrEmpty(valueText))
                return null;

            var channelEntry = channels.Keys.FirstOrDefault(key => valueText.Contains(key));

            if (string.IsNullOrEmpty(channelEntry))
                return null;
            else if (includeGuid)
                return $"{channels[channelEntry]}: ({channelEntry})";
            else
                return channels[channelEntry];
        }


        public static AddinApplicationInfo Addin => OfficeAddinConnection.AddinEntryAssembly;

        /// <summary>
        ///     Returns the maximum size available for the supplied TaskPane
        /// </summary>
        /// <param name="pane">The Pane to calculate the size for.</param>
        public static void GetMaximumDockedSize(Window wordWindow, out int width, out int height)
        {
            var w = new WindowInfo(WordWindowHelper.GetMdiClientWindowHandle(wordWindow));
            var sz = w.Size;
            width = sz.Width;
            height = sz.Height;
        }


        //internal static WordAddinApplication WordAddinApplication { get; set; }

        /// <summary>
        ///     Defines the TraceSwitch Level to use for altering the Trace output from the WordExtensions Assembly
        /// </summary>
        /// <value>A TraceLevel Value.</value>
        /// <returns>One of the TraceLevel values.</returns>

        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public static TraceLevel TraceLevel
        {
            get => StaticTraceSwitch.TraceSwitch.Level;
            set => StaticTraceSwitch.TraceSwitch.Level = value;
        }

        /// <summary>
        /// Returns true if the current version of Word is Dpi Aware.
        /// Currently this is Word 2016 version 1803 (build 9126.2116) or greater. It is also dependent on the OS being at least Windows 10, Anniversary addition.
        /// </summary>

        public static bool IsWordVersionDpiAware()
        {
            if (!OSVersionHelper.IsWindows10AnniversaryAdditionOrGreater())
                return false;

            if (!Version.TryParse(WordApplication.Build, out var version))
                return false;
            //Improvements for Word, PowerPoint, and Visio were released on March 27, 2018 for customers running version 1803 (build 9126.2116) or greater.
            // https://support.office.com/en-us/article/office-apps-appear-the-wrong-size-or-blurry-on-external-monitors-bc9f7279-4e42-4b15-a949-46ab8bcfe44f
            return version.Major > 16 || version.Major == 16 && version.Build >= 9126;
        }

        private static OfficeVersion _wordVersion;

        /// <summary>
        ///     Returns an enum representing the major version number of Microsoft
        /// </summary>

        /// <returns>One of the defined WordVersion values.</returns>

        public static OfficeVersion WordVersion => WordApplication.VersionLJ();

        /// <summary>
        ///     Returns the current Theme applied to Office 2007 and above.
        /// </summary>

        /// <returns>One of the OfficeTheme values.</returns>

        public static OfficeUITheme Theme
        {
            get
            {
                if (WordWindowHelper == null)
                {
                    return OfficeUITheme.None;
                }

                return WordWindowHelper.Theme;
            }
        }




        /// <summary>
        ///     Returns the path of the assembly that this class is executing from
        /// </summary>
        public static string ExecutingPath
        {
            get
            {
                if (string.IsNullOrEmpty(_addinPath))
                {
                    // ReSharper disable once AssignNullToNotNullAttribute
                    var assemblyUri = new Uri(Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase));
                    _addinPath = assemblyUri.LocalPath;
                }

                return _addinPath;
            }
        }

        /// <summary>
        ///     Returns a IWin32Window interface that represents the Application window belonging to the Window passed.
        /// </summary>
        /// <param name="window">The Window you want to return the Word application window handle for</param>
        /// <returns>A IWin32Window interface.</returns>
        /// <remarks>
        ///     In The versions of Word that has ShowWindowsInTaskBar set to False all the IWin32Window window handles will be the
        ///     same.
        ///     If the window param is null an exception is thrown.
        /// </remarks>
        [CLSCompliant(isCompliant: false)]
        public static IntPtr GetWordApplicationWindow(Window window = null)
        {
            if (WordWindowHelper == null)
                return IntPtr.Zero;

            if (window == null)
                window = WordApplication.ActiveWindowLJ();

            return WordWindowHelper.GetApplicationHandle(window);
        }

        public static IntPtr GetWordApplicationWindowObj(object window = null) =>
            GetWordApplicationWindow((Window)window);


        /// <summary>
        ///     Returns if the supplied wordWindow belongs to the active top level Word window.
        /// </summary>
        /// <param name="wordWindow"></param>

        /// <remarks>
        ///     This is not the Same as Window.IsActive. This method returns if it is the Active Window in Windows where as
        ///     Window.IsActive returns if it is the active window in  wordWindow is of type object so UI code codes not need a
        ///     reference to the Word Assemblies.
        /// </remarks>
        public static bool IsWordWindowActive(Window wordWindow)
        {
            Check.NotNull(wordWindow, "wordWindow");
            if (wordWindow.Active)
            {
                return new WindowInfo(GetWordApplicationWindow(wordWindow)).IsActiveWindow;
            }

            return false;
        }


        // private members


        private static void AttachEventHandlers()
        {
            if (WordWindowHelper == null)
            {
                return;
            }

            DetachEventHandlers();
 
            WordWindowHelper.WordThemeChanged += Singleton.WordThemeChangedHandler;
            WordWindowHelper.WindowClosed += Singleton.WordWindowClosedHandler;
 
            WordWindowHelper.WindowActivateChanged += Singleton.WordWindowActivateChangedHandler;
 
            if (WordApplication != null)
            {
                ((ApplicationEvents4_Event)WordApplication).Quit += Singleton.WordApplicationQuitHandler;
                WordApplication.DocumentChange += Singleton.WordDocumentChangeHandler;
                _wordEventsAttached = true; //Word handlers need a flag
            }
        }


        private static void DetachEventHandlers()
        {
            if (_wordEventsAttached && WordApplication != null)
            {
                ((ApplicationEvents4_Event)WordApplication).Quit -= Singleton.WordApplicationQuitHandler;
                WordApplication.DocumentChange -= Singleton.WordDocumentChangeHandler;
                _wordEventsAttached = false;
            }

            if (WordWindowHelper != null)
            {
                WordWindowHelper.WordThemeChanged -= Singleton.WordThemeChangedHandler;
                WordWindowHelper.WindowClosed -= Singleton.WordWindowClosedHandler;
                WordWindowHelper.WindowActivateChanged -= Singleton.WordWindowActivateChangedHandler;
            }
        }


        internal static bool Quitting { get; private set; }

        public static IWordSelectionChange SelectionChange
        {
	        get => _wordSelectionChange;

            set
	        {
                if (_wordSelectionChange != null)
                    throw new InvalidOperationException("Instance already set");
                Check.NotNull(value, nameof(value));

                _wordSelectionChange = value;
                _wordSelectionChange.SelectionChanged -= Singleton.SelectionChangedHandler;
                _wordSelectionChange.SelectionChanged += Singleton.SelectionChangedHandler;
            }
        }

        private void SelectionChangedHandler(object sender, WordSelectionEventArgs e)
        {
	        SelectionChanged?.Invoke(this, e);
        }


        private void WordApplicationQuitHandler()
        {
	        Quitting = true;
	        WordApplicationClosing?.Invoke(Singleton, EventArgs.Empty);
        }


        public static WordWindowHelper WordWindowHelper { get; private set; }


        private void WordDocumentChangeHandler()
        {
            try
            {
                if (Singleton._wordUndoRecord != null)
                {
                    UndoRecord.EnsureRecordingCustomRecord();
                }
                WordDocumentChange?.Invoke(Singleton, EventArgs.Empty);
            }
            catch (InvalidComObjectException)
            {
                Singleton.Dispose();
            }
        }


        /// <summary>
        ///     Fires the WordThemeChanged Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void WordThemeChangedHandler(object sender, EventArgs e)
        {
            WordThemeChanged?.Invoke(this, e);
        }


        /// <summary>
        ///     Fires the WordWindowClosed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void WordWindowClosedHandler(object sender, WordWindowEventArgs e)
        {
            WordWindowClosed?.Invoke(this, e);
        }


        private void WordWindowActivateChangedHandler(object sender, WordWindowActivateEventArgs e)
        {
            WordWindowActivateChanged?.Invoke(this, e);
            
        }

        public static bool HasVisibleDocuments()
        {
            var windows = WordApplication.Windows;
            for (var i = 1; i <= windows.Count; i++)
            {
                if (windows[i].Visible)
                {
                    return true;
                }
            }

            return false;
        }


        public static int CountOfVisibleWordWindows()
        {
            var windows = WordApplication.Windows;
            var visibleWordWindows = 0;
            for (var i = 1; i <= windows.Count; i++)
            {
                if (windows[i].Visible)
                {
                    visibleWordWindows++;
                }
            }

            return visibleWordWindows;
        }


        public static bool WindowContainsFocus(Window window)
        {
            Check.NotNull(window, "window");
            var windowInfo = new WindowInfo(WordWindowHelper.GetDocumentWindowHandle(window));

            if (windowInfo.IsValid)
            {
                var focus = WindowInfo.FromFocus();
                if (focus == windowInfo)
                {
                    return true;
                }

                do
                {
                    focus = focus.Parent;
                    if (!focus.IsValid)
                    {
                        return false;
                    }

                    if (windowInfo == focus)
                    {
                        return true;
                    }
                } while (true);
            }

            return false;
        }


        public static void SetFocusToWordWindow(Window window, bool activateWordWindow = true)
        {
            Check.NotNull(window, "window");
            if (activateWordWindow)
                window.ActivateLJ();

            var windowInfo = new WindowInfo(WordWindowHelper.GetDocumentWindowHandle(window));
            if (windowInfo.IsValid)
            {
                windowInfo.SetFocus();
            }
        }


        public static bool IsFocusInWordWindow()
        {
            IsFocusInWordWindow(window: null);

            return false;
        }


        public static bool IsFocusInWordWindow(Window window, bool ignoreModelessWindows = false)
        {
            Check.NotNull(window, "window");
            //window.ActivateLJ()
            var windowInfo = WindowInfo.FromFocus();
            if (windowInfo.IsValid)
            {
                if (windowInfo.ClassAtom != WordWindowHelper.WwGWindowClass)
                {
                    if (ignoreModelessWindows && window.Active)
                    {
                        if (windowInfo.Owner != null &&
                            windowInfo.Owner.ClassAtom == WordWindowHelper.OpusAppWindowClass &&
                            windowInfo.Owner.Visible == false)
                        {
                            return true;
                        }
                    }

                    return false;
                }

                if (window != null)
                {
                    var windowInfo2 = new WindowInfo(WordWindowHelper.GetDocumentWindowHandle(window));
                    return windowInfo2 == windowInfo.Parent;
                }

                return true;
            }

            return false;
        }


        public static bool IsFocusInWordModelessDialog()
        {
            var windowInfo = WindowInfo.FromFocus();
            if (windowInfo.IsValid)
            {
                if (windowInfo.Owner != null && windowInfo.Owner.ClassAtom == WordWindowHelper.OpusAppWindowClass &&
                    windowInfo.Owner.Visible == false)
                {
                    return true;
                }
            }

            return false;
        }


        public static OfficeShortcutKeyCollection ShortcutKeys
        {
            get
            {
                if (_shortcutKeys == null)
                {
                    _shortcutKeys = new OfficeShortcutKeyCollection();
                }

                return _shortcutKeys;
            }
        }


        internal static void RaiseShortcutKeyPressed(OfficeShortcutKeyPressedEventArgs e)
        {
            ShortcutKeyPressed?.Invoke(Singleton, e);
        }


        public static string CommandLine
        {
            get
            {
                var ptr = GetCommandLine();
                if (ptr != IntPtr.Zero)
                {
                    return Marshal.PtrToStringAuto(ptr);
                }

                return null;
            }
        }


        public static Func<SynchronizationContext> SynchronizationContextFactory { get; set; }

        //<DebuggerStepThrough()>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public static void BeginInvoke<T>(Action<T> method, object state)
        {
	        SynchronizationContextFactory?.Invoke()?.Post((o) => method((T)o), state);
        }


        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public static void BeginInvoke(Action method)
        {
	        SynchronizationContextFactory?.Invoke()?.Post((o) =>
            {
                try
                {
                    method();
                }
                catch (Exception ex)
                {
                    throw;
                }
            },null);
        }

        /// <summary>
        ///     Locks Words ScreenUpdating.
        /// </summary>

        /// <remarks>
        ///     This method uses reference counting so ScreenUpdating is not turned back on until the same number of
        ///     UnLockScreenUpdating has been called or UnLockScreenUpdating is called with the reset set to true. It is
        ///     recommended that Locking and unlocking is done using a Try Finally Block to guarantee the LockScreenUpdating are
        ///     always balanced with the same number of UnLockScreenUpdating calls.
        /// </remarks>
        public static bool LockScreenUpdating()
        {
            if (_wordScreenUpdateLock.Lock())
            {
                if (AppDiagnostics.GetOption(AppDiagnosticOptions.SuppressScreenLocking) == false)
                {
                    WordApplication.ScreenUpdating = false;
                }

                return true;
            }

            return false;
        }


        /// <summary>
        ///     UnLocks Words ScreenUpdating.
        /// </summary>
        /// <param name="reset">
        ///     Resets the screen locking, and turns painting back on. This member should only be used in rare
        ///     cases, such as unhandled exceptions.
        /// </param>

        /// <remarks>
        ///     ScreenUpdating is not turned back on until UnLockScreenUpdating has been called the same number of times as
        ///     the LockScreenUpdating call or UnLockScreenUpdating is called with the reset set to true.
        /// </remarks>
        public static bool UnLockScreenUpdating(bool reset = false)
        {
            if (reset)
            {
                WordApplication.ScreenUpdating = true;
                WordApplication.ScreenRefresh();
                _wordScreenUpdateLock.Reset();
                return true;
            }

            if (_wordScreenUpdateLock.Unlock())
            {
                if (AppDiagnostics.GetOption(AppDiagnosticOptions.SuppressScreenLocking) == false)
                {
                    WordApplication.ScreenUpdating = true;
                    WordApplication.ScreenRefresh();
                }

                return true;
            }

            return false;
        }


        public static void SuspendScreenUpdating()
        {
            if (_wordScreenUpdateLock.Locked)
            {
                WordApplication.ScreenUpdating = true;
            }
        }


        public static void ResumeScreenUpdating()
        {
            if (_wordScreenUpdateLock.Locked)
            {
                WordApplication.ScreenUpdating = false;
            }
        }


        public static bool IsScreenUpdatingLocked => _wordScreenUpdateLock.Locked;


        /// <summary>
        ///     Indicates if the Word application is in the Backstage view.
        ///     Set when the Backstage Ribbon xml callback onShow if fired.
        /// </summary>
        public static bool InBackstage { get; internal set; }


        /// <summary>
        ///     Stores the focus, to be restored when the EndFocusChangeAction is called
        /// </summary>
        /// <param name="activateActiveWordWindow"></param>

        /// <remarks>
        ///     The BeginFocusChangeAction,EndFocusChangeAction should be placed in a Try Finally block to ensure that the
        ///     EndFocusChangeAction is called.
        /// </remarks>
        public static bool BeginFocusChangeAction(bool activateActiveWordWindow)
        {
            if (_focusChangeActionLock.Lock())
            {
                _priorFocus = WindowInfo.FromFocus();
                if (activateActiveWordWindow)
                {
                    var activeWindow = WordApplication.ActiveWindowLJ();
                    if (activeWindow != null)
                    {
                        if (IsWordWindowActive(activeWindow))
                        {
                            activeWindow.ActivateLJ();
                        }
                    }
                }

 
                return true;
            }

            return false;
        }
        public static WindowInfo PriorFocus => _priorFocus;

        public static bool InFocusChangeAction => _focusChangeActionLock.Locked;
 
        public static bool EndFocusChangeAction()
        {
            if (_focusChangeActionLock.Unlock())
            {
                var priorFocus = _priorFocus;
                _priorFocus = null;
                if (priorFocus != null && priorFocus.IsValid)
                {
                    var rootFocus = priorFocus.Root;
                    if (rootFocus != null)
                    {
                        rootFocus.Activate();

                    }

                    priorFocus.SetFocus();
                }

                return true;
            }

            return false;
        }


        /// <summary>
        ///     Executes Word Actions in the provide action delegate that change or require the focus to be in the word window when
        ///     run. After execution the focus is restored to it's original location
        /// </summary>
        /// <param name="action">The delegate that executes the word actions</param>
        /// <param name="activateActiveWordWindow">Activates the Active Word Window prior to invoking the action delegate</param>

        public static void ExecuteFocusChangeAction(bool activateActiveWordWindow, Action action)
        {
            Check.NotNull(action, "action");
            BeginFocusChangeAction(activateActiveWordWindow);
            try
            {
                action.Invoke();
            }
            finally
            {
                EndFocusChangeAction();
            }
        }
 

        public static WordUndoRecord UndoRecord
        {
            get
            {
                if (Singleton == null)
                {
                    return null;
                }

                if (Singleton._wordUndoRecord == null)
                {
                    Singleton._wordUndoRecord = new WordUndoRecord();
                }

                return Singleton._wordUndoRecord;
            }
        }


        public static SuppressWordEvents SuppressEvents
        {
            get
            {
                if (Singleton == null)
                {
                    return null;
                }

                return Singleton._suppressWordAppEvents ?? (Singleton._suppressWordAppEvents = new SuppressWordEvents());
            }
        }


        public static void InvalidateRibbon()
        {
            Addin?.Connection.InvalidateRibbon();
        }

        public static void InvalidateRibbonControl(string ribbonControlId)
        {
            Addin?.Connection.InvalidateRibbonControl(ribbonControlId);
        }

        [CLSCompliant(false)]
        public static Application GetOrStartNewWordApplication(bool visible = true, bool addDocument = true)
        {
            return ((Application)null).GetOrStartNewWordApplication(visible, addDocument);
        }

        internal void OnAddinDisconnected()
        {
            AddinDisconnected?.Invoke(Singleton, EventArgs.Empty);
        }


        public static bool IsX64()
        {
            if (Addin != null)
                return IntPtr.Size == 8; //We are loaded inside word so we can just check the pointer size.

            //Does not work if registry redirection is being used
            //var program = ApplicationHelper.GetAssociatedApplicationFromProgramId("Word.Document");

            if (!GetPath(Registry.ClassesRoot, out var wordExe))
                if (!GetPath(Registry.LocalMachine, out wordExe))
                    if (!GetPath(Registry.CurrentUser, out wordExe))
                        return false;

            return ApplicationHelper.GetNativeBitness(wordExe) == Bitness.x64;

            bool GetPath(RegistryKey hive, out string wordAppExe)
            {
                wordAppExe = hive.GetKeyValue(@"Software\Classes\CLSID\{F4754C9B-64F5-4B40-8AF4-679732AC0607}\LocalServer32", "") as string;
                if (wordAppExe != null && File.Exists(wordAppExe))
                    return true;
                wordAppExe = hive.GetKeyValue(@"Software\Classes\WOW6432Node\CLSID\{F4754C9B-64F5-4B40-8AF4-679732AC0607}\LocalServer32", "") as string;

                return wordAppExe != null && File.Exists(wordAppExe);
            }
        }


        // IDisposible (remember to detach those event handlers!)


        private bool _disposedValue; // To detect redundant calls

        // IDisposable
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions")]
        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters",
            MessageId = "System.Windows.Forms.MessageBox.Show(System.String)")]
        private void Dispose(bool disposing)
        {
            //Debug.WriteLine("Dispose: " & Me.GetType.Name)

            if (!_disposedValue)
            {
                if (disposing)
                {

                    try
                    {
                        DetachEventHandlers();

                        if (_wordSelectionChange != null)
                        {
                            _wordSelectionChange.Options = SelectionChangeOption.None;
                            _wordSelectionChange = null;
                        }

                        if (WordWindowHelper != null)
                        {
                            WordWindowHelper.Dispose();
                            WordWindowHelper = null;
                        }

                        if (_shortcutKeys != null)
                        {
                            _shortcutKeys.Clear();
                            _shortcutKeys = null;
                        }

                        if (_font != null)
                        {
                            _font.Dispose();
                            _font = null;
                        }

                        Addin?.Reset();

                        Singleton = null;
                        //_dialogManager = null;
                        WordApplication = null;

                        //MessageBox.Show("Word Extensions Disposed")
                    }
                    catch (Exception ex)
                    {
                        //If we get an exception in here and we do not catch it then we may crash Word
                        if (Debugger.IsAttached)
                            Debug.Assert(false, @"*** DISPOSE EXCEPTION IN W.E." + Environment.NewLine + ex);
                    }
                }

                // TODO: free your own state (unmanaged objects).
                // set large fields to null.
            }

            _disposedValue = true;
        }

        /// <summary>
        ///     Disposes of the resources (other than memory) used by the AddinConnection object.
        /// </summary>

        public void Dispose()
        {
            // Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        ~WordExtensions()
        {
#if (TRACK_DISPOSED)
            LevitJames.Core.Diagnostics.DisposeTracker.Add(_disposedSource);
#endif
            Dispose(disposing: false);
        }
         
    }
}