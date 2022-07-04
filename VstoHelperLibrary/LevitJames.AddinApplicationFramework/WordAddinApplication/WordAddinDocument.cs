using LevitJames.Core;
using LevitJames.MSOffice;
using Microsoft.Office.Interop.Word;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace LevitJames.AddinApplicationFramework
{
    public enum SaveOption
    {
        SaveNoUserPrompt,
        DontSave,
        SaveWithUserPrompt
    }

    public abstract class WordAddinDocument : IDisposable, IWordDocumentProvider
    {
        private ServiceProvider _serviceProvider;
        private Window _workingWindow;
        private TemporaryFile _tempFile;

        public event EventHandler SelectionChanged;
        public event EventHandler<CancelEventArgs> DocumentClosing;
        public event EventHandler DocumentClosed;
        public event EventHandler<WordAddinDocumentBeforeSaveEventArgs> DocumentBeforeSave;
        public event EventHandler<WordAddinDocumentWindowEventArgs> WindowActivated;
        public event EventHandler<WordAddinDocumentWindowEventArgs> WindowDeactivated;
        public event EventHandler WindowViewTypeChanged;
        public event EventHandler EditableChanged;
        public event EventHandler ProtectedChanged;
#if (TRACK_DISPOSED)
        private readonly string _disposedSource;
#endif

        protected WordAddinDocument()
        {
#if (TRACK_DISPOSED)
            _disposedSource = LevitJames.Core.Diagnostics.DisposeTracker.GetSource();
#endif
        }

        protected internal virtual void Initialize(object addinApp, Document wordDocument)
        {
            Check.NotNull(wordDocument, "wordDocument");

            WordDocument = wordDocument;
            IsEditable = wordDocument.LJIsEditable() == WordBoolean.True;
        }


        /// <summary>
        ///     Returns the name of the Word document This AppDocument is working on
        /// </summary>
        public string Name => WordDocument?.Name;


        /// <summary>
        ///     Returns the filename of the Word document This AppDocument is working on
        /// </summary>
        public string FileName => WordDocument?.FullNameLJ();


        public object Parent { get; internal set; }


        private IAddinApplication ParentInternal => (IAddinApplication)Parent;

        internal void SetTemporaryFile(TemporaryFile tempFile)
        {
            _tempFile = tempFile;
        }


        public string OriginalFileName => _tempFile == null ? FileName : _tempFile.OriginalFileName;


        /// <summary>
        ///     Returns Sets the window that the Add-in is currently working in.
        /// </summary>



        public Window WorkingWordWindow
        {
            get => _workingWindow;
            set
            {
                if (_workingWindow == null && value == null)
                    return;

                _workingWindow = value;
                WdViewType viewType = 0;
                if (_workingWindow != null)
                {
                    viewType = _workingWindow.View.Type;
                }
                else if (WordDocument != null && WordExtensions.WordApplication.IsObjectValid[WordDocument])
                {
                    if (WordDocument.ActiveWindow != null)
                    {
                        viewType = WordDocument.ActiveWindow.View.Type;
                    }
                    else
                    {
                        viewType = 0;
                    }
                }

                if (viewType != WorkingWindowViewType)
                {
                    WorkingWindowViewType = viewType;
                    OnWindowViewTypeChanged(EventArgs.Empty);
                }
            }
        }


        /// <summary>
        ///     Returns The ViewType of the Working WorkingWordWindow
        /// </summary>


        /// <remarks>This value is monitored via the OnIdle method and raises the WindowViewTypeChanged when it changes</remarks>
        public WdViewType WorkingWindowViewType { get; private set; }


        internal void UpdateDocument()
        {
            var raiseEditableChanged = false;

            if (WordDocument == null) //May be null
            {
                return;
            }

            if (WordExtensions.WordApplication == null)
            {
                return;
            }

            if (WordExtensions.WordApplication.IsObjectValid[WordDocument] == false)
            {
                WordDocument = null;
                return;
            }

            var editable = WordDocument.LJIsEditable();
            if (editable != WordBoolean.Unknown)
            {
                var isEditable = editable == WordBoolean.True;
                if (IsEditable != isEditable)
                {
                    IsEditable = isEditable;
                    raiseEditableChanged = true;
                }
            }

            var raiseWindowViewTypeChanged = ChangeWindowViewTypeChangedIfRequired();

            // Only raise events after all the states have been checked.
            // This is to allow all the events to be able to get the correct state values.
            //If this is the first time then we raise always raise the events.

            if (raiseEditableChanged)
            {
                OnEditableChanged(EventArgs.Empty);
            }

            if (raiseWindowViewTypeChanged)
            {
                OnWindowViewTypeChanged(EventArgs.Empty);
            }
        }


        private bool ChangeWindowViewTypeChangedIfRequired()
        {
            WdViewType viewType;
            if (_workingWindow == null || !WordDocument.Application.IsObjectValid[_workingWindow])
            {
                _workingWindow = null;
                if (WordDocument.ActiveWindow != null)
                {
                    viewType = WordDocument.ActiveWindow.View.Type;
                }
                else
                {
                    viewType = 0;
                }
            }
            else
            {
                viewType = _workingWindow.View.Type;
            }

            if (viewType != WorkingWindowViewType)
            {
                WorkingWindowViewType = viewType;
                return true;
            }

            return false;
        }


        /// <summary>
        ///     Raises the EditableChanged Event when the Editable state has changed
        /// </summary>
        /// <param name="e"></param>

        protected virtual void OnEditableChanged(EventArgs e)
        {
            EditableChanged?.Invoke(this, e);
        }


        /// <summary>
        ///     Raises the ProtectedChanged Event when the Protected state has changed
        /// </summary>
        /// <param name="e"></param>

        protected virtual void OnProtectedChanged(EventArgs e)
        {
            ProtectedChanged?.Invoke(this, e);
        }


        protected internal virtual void OnDocumentBeforeSave(WordAddinDocumentBeforeSaveEventArgs e)
        {
            DocumentBeforeSave?.Invoke(this, e);
        }


        protected internal virtual void OnAdded() { }


        /// <summary>
        ///     Raises the WindowViewTypeChanged Event when the WorkingWordWindow.View.Type changes
        /// </summary>
        /// <param name="e"></param>

        protected virtual void OnWindowViewTypeChanged(EventArgs e)
        {
            WindowViewTypeChanged?.Invoke(this, e);
        }


        // ''' <summary>
        // ''' Raises the CompatibilityChanged Event when a Word Document's Compatibility has changed
        // ''' </summary>
        // ''' <param name="e"></param>
        // ''' <remarks>Compatibility changes depending on the type of document being edited, older documents types such as .doc,  .dot and .rtf are considered legacy documents.</remarks>
        //Protected Overridable Sub OnCompatibilityChanged(e As EventArgs)
        //	RaiseEvent CompatibilityChanged(Me, e)
        //End Sub


        public virtual void Close()
        {
            Close(SaveOption.SaveWithUserPrompt);
        }

        public virtual void Close(SaveOption option)
        {
            var parent = ParentInternal;
            switch (option)
            {
                case SaveOption.SaveNoUserPrompt:
                    if (!WordDocument.Saved)
                        WordDocument.Save();
                    break;
                case SaveOption.DontSave:
                    WordDocument.CloseLJ(dontPromptToSave: true);
                    break;
                case SaveOption.SaveWithUserPrompt:
                    WordDocument.Close();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(option), option, message: null);
            }

            parent?.EnsureActiveDocument();
        }


        public bool IsActive
        {
            get
            {
                if (WordDocument == null)
                {
                    return false;
                }

                return WordDocument.Application.ActiveDocumentLJ() == WordDocument;
            }
        }

        /// <summary>
        ///     Calls WordDocument.Activate and additionally places the focus to the document window .
        /// </summary>

        public void Activate()
        {
            Activate(setKeyboardFocus: true);
        }

        /// <summary>
        ///     Calls WordDocument.Activate.
        /// </summary>
        /// <param name="setKeyboardFocus">
        ///     If true the keyboard focus is additionally set to place the focus to the document
        ///     window.
        /// </param>

        public void Activate(bool setKeyboardFocus)
        {
            if (WordDocument == null)
            {
                return;
            }

            WordDocument.Application.Activate();
            WordDocument.Activate();
            if (setKeyboardFocus)
            {
                WordExtensions.SetFocusToWordWindow(WordDocument.ActiveWindow);
            }
        }

        public Document WordDocument { get; private set; }
        Document IWordDocumentProvider.Document => WordDocument;

        public bool IsEditable { get; set; }

        /// <summary>
        ///     Raises the DocumentClosing event for this WordAddinDocument
        /// </summary>
        /// <param name="e">Any EventArgs to pass to the event.</param>

        protected internal virtual void OnDocumentClosing(CancelEventArgs e)
        {
            DocumentClosing?.Invoke(this, e);
        }

        /// <summary>
        ///     Raises the WindowActivated event when a window belonging to the this WordAddinDocument is activated
        /// </summary>
        /// <param name="e">Any EventArgs to pass to the event.</param>

        protected internal virtual void OnWindowActivated(WordAddinDocumentWindowEventArgs e)
        {
            WindowActivated?.Invoke(this, e);
        }

        /// <summary>
        ///     Raises the WindowDeactivated event when a window belonging to the this WordAddinDocument is de-activated
        /// </summary>
        /// <param name="e">Any EventArgs to pass to the event.</param>

        protected internal virtual void OnWindowDeactivated(WordAddinDocumentWindowEventArgs e)
        {

            WindowDeactivated?.Invoke(this, e);
        }

        /// <summary>
        ///     Raises the DocumentClosed event for this WordAddinDocument
        /// </summary>

        internal void OnDocumentClosed()
        {
            WordDocument = null;
            OnDocumentClosed(EventArgs.Empty);
            Dispose();
        }

        /// <summary>
        ///     Raises the DocumentClosed event for this WordAddinDocument
        /// </summary>
        /// <param name="e">Any EventArgs to pass to the event.</param>

        protected virtual void OnDocumentClosed(EventArgs e)
        {
            DocumentClosed?.Invoke(this, e);
        }

        /// <summary>
        ///     Raises the SelectionChanged event for this WordAddinDocument
        /// </summary>
        /// <param name="e">Any EventArgs to pass to the event.</param>

        protected internal virtual void OnWordDocumentSelectionChanged(EventArgs e)
        {
            SelectionChanged?.Invoke(this, e);
        }

        /// <summary>
        ///     Called by the WordAddinApplication when an unhandled exception occurs while this WordAddinDocument is active.
        /// </summary>
        /// <param name="exception"></param>
        /// <returns>true if this method handled the exception;false otherwise. The default is false.</returns>

        protected internal virtual bool OnThreadException(Exception exception)
        {
            return false;
        }

        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public ServiceProvider ServiceProvider => _serviceProvider ?? (_serviceProvider = new ServiceProvider());

        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public TClass GetService<TClass>() where TClass : class
        {
            return _serviceProvider?.GetService<TClass>() ??
                   (TClass)((IServiceProvider)ParentInternal)?.GetService(typeof(TClass));
        }

        /// <summary></summary>
        ~WordAddinDocument()
        {
#if (TRACK_DISPOSED)
            LevitJames.Core.Diagnostics.DisposeTracker.Add(_disposedSource);
#endif
            Dispose(disposing: false);
        }

        protected virtual void Dispose(bool disposing)
        {
            //Debug.WriteLine("Dispose: " & Me.GetType.Name)

            if (_workingWindow != null)
            {
                Marshal.ReleaseComObject(_workingWindow);
                _workingWindow = null;
            }

            if (WordDocument != null)
            {
                Marshal.ReleaseComObject(WordDocument);
                WordDocument = null;
            }

            if (_tempFile != null)
            {
                _tempFile.Dispose();
                _tempFile = null;
            }

            _serviceProvider = null;
            Parent = null;
            if (disposing)
                GC.SuppressFinalize(this);
        }


        public void Dispose()
        {
            Dispose(disposing: true);
        }
    }
}
