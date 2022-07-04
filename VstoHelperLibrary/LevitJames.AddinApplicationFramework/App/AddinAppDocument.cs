// © Copyright 2018 Levit & James, Inc.

using System;
using System.ComponentModel;
using System.Runtime.Serialization;
using LevitJames.Core;
using LevitJames.MSOffice;
using LevitJames.MSOffice.MSWord;
using Microsoft.Office.Interop.Word;

namespace LevitJames.AddinApplicationFramework
{
    /// <summary>
    ///     An abstract class which wraps a Word Document, providing additional functionality.
    /// </summary>
    public abstract class AddinAppDocument : WordAddinDocument, IAddinAppProvider, ISerializable, IAddinAppDirty
    {
        internal bool FillProxy;

        public event EventHandler DocumentStateChanged;
        public event EventHandler TransactionStarted;

        public event EventHandler<TransactionCompletedEventArgs> TransactionCompleted;
        public event EventHandler UndoCompleted;

#if (TRACK_DISPOSED)
        private readonly string _disposedSource;
#endif

        protected AddinAppDocument()
        {
            Session = new AddinAppSession();
#if (TRACK_DISPOSED)
            _disposedSource = LevitJames.Core.Diagnostics.DisposeTracker.GetSource();
#endif
        }


        protected AddinAppDocument(SerializationInfo info, StreamingContext context)
        {
            var doc = FillProxy ? this : (AddinAppDocument) context.Context;
#pragma warning disable CA2214 // Do not call overridable methods in constructors
            doc.OnDeserialize(AppSerializationState.OnDeserialize(doc, info, context));
#pragma warning restore CA2214 // Do not call overridable methods in constructors
        }

        void ISerializable.GetObjectData(SerializationInfo info, StreamingContext context)
        {
            OnSerialize(AppSerializationState.OnSerialize(info, context));
        }


        public virtual bool AutoDirty => true;


        protected internal virtual void OnDeserialize(AppSerializationState state) { }


        protected virtual void OnSerialize(AppSerializationState state) { }


        public IAddinApplication App => AppInternal;
        internal IAddinApplicationInternal AppInternal { get; private set; }

        protected internal override void Initialize(object app, Document wordDocument)
        {
            base.Initialize(app, wordDocument);
            AppInternal = (IAddinApplicationInternal) app;
            ((IAddinAppProviderInternal) Session).Initialize(AppInternal, this);
            Store = CreateStore();
            WordChangeSets = new WordSettingsChangeSetController(wordDocument);

            CheckIsNew();

            OnInitialize();
        }


        protected virtual void OnInitialize() { }


        /// <summary>
        ///     Returns A Session instance containing information about the current session state of the document.
        /// </summary>
        public AddinAppSession Session { get; }

        /// <summary>
        ///     Returns if this document contains data for the application.
        /// </summary>
        
        public bool IsAppDocument => Store.IsAppDocument;


        public bool IsNew { get; private set; }


        public bool CheckIsNew()
        {
            IsNew = WordDocument.Range().End == 1 &&
                    WordDocument.Name.StartsWith("document", StringComparison.OrdinalIgnoreCase);

            return IsNew;
        }


        //public AddinAppTransactionHistory History { get; private set; }

        private bool _sessionHistoryWritten;

        private AddinAppTransactionHistory _history;

        public AddinAppTransactionHistory History
        {
            get => _history ?? (_history ?? Store.GetHistory(() => History));
            // ReSharper disable once UnusedMember.Local
            private set
            {
                if (_history != null)
                    _sessionHistoryWritten = _history.SessionHeaderWritten;

                _history = value;
                if (_history != null)
                {
                    ((IAddinAppProviderInternal) History).Initialize(AppInternal, this);
                    _history.SessionHeaderWritten = _sessionHistoryWritten;
                }
            }
        }


        public WordSettingsChangeSetController WordChangeSets { get; private set; }


        internal void EnterSession()
        {
            if (Session.InSession)
                throw new InvalidOperationException("Already in session");

            Session.InSession = true;
            Store.OnEnterSession();

            OnEnterSession();
        }


        /// <summary>
        ///     Override to add any additional code that may be required at the end of entering a Session.
        /// </summary>
        protected virtual void OnEnterSession() { }


        internal void ExitAllSessions()
        {
            if (!Session.InSession)
                return;

            App.Tracer.TraceInformation("ExitAllSessions in");

            Session.CloseAllSessions();

            Store.OnExitSession();

            OnExitSession();
            Session.InSession = false;

            App.Tracer.TraceInformation("ExitAllSessions out");
        }


        /// <summary>
        ///     Override to add any additional code that may be required before finally exiting from a Session.
        /// </summary>
        protected virtual void OnExitSession()
        {
            //Default implementation does nothing.
        }

        /// <summary>
        /// Clear any object model instances. Called after OnExitSession. Unlike OnExitSession this method is called even when 
        /// </summary>
        protected internal virtual void OnClearModel()
        {
            Store?.Instances.Clear();
        }

        public bool GetExclusiveEventsFlag() => Store.GetBool(AddinAppDocumentStorage.DisableWordEventsVariableName);


        /// <summary>Adds or removes the LevitJames.LJRequiresExclusiveWordEvents variable for use by other Word Addins/COMAddins.</summary>
        /// <param name="add">If true, variable is removed</param>
        public void SetExclusiveEventsFlag(bool add)
        {
            Store.SetBool(AddinAppDocumentStorage.DisableWordEventsVariableName, add);
        }


        protected internal override void OnDocumentClosing(CancelEventArgs e)
        {
            base.OnDocumentClosing(e);
            if (e.Cancel)
                return;

            // Allow the Active transaction to Close the document if it is inside InOnInitiate
            // If its a Modeless transaction InOnInitiate will be false.
            if (Session.ActiveTransaction?.InOnInitiate == true)
                return;

            App.ViewService.OnOwnerClosing(this);

            SetExclusiveEventsFlag(false);

            if (Session.InSession)
                AppInternal.TransactionManager.EndEditSession(this);

            ExitAllSessions();
        }


        //private void LoadDocumentState()
        //{
        //    InitializeStorageManager();
        //    StorageManager.LoadDocumentState();

        //    CloseStorageManager();
        //}


        internal bool InitialEnterSessionDone { get; set; }


        protected internal override void OnAdded()
        {
            // KDP NOTE: The ToggleViewType method was added 04/16/2013 for a Word 2013
            // bug for read-only documents running a Range.InsertAfter API call. As of
            // 07/30/2019, this bug is no longer evident, and so we are removing the 
            // call to ToggleViewType as it is interfering with the document window painting.
            // C.f., TSWA 7114.

            // ToggleViewType();
            CheckDocumentStateChanged();
            DocumentRecovery();
            UserActionManager.UpdateDirty = true;
        }


        protected override void OnDocumentClosed(EventArgs e)
        {
            base.OnDocumentClosed(e);
            UserActionManager.UpdateDirty = true;
        }


        private void DocumentRecovery()
        {
            if (Store.RecoveryRequired)
                App.CreateDocumentRecovery()?.Recover(this);
        }


        private void ToggleViewType()
        {
            // Jiggling the view type on Read-Only documents gets around the problem in Word 2013:
            // "this method or property is not available because this command is not available for reading"
            // when running a Range.InsertAfter call. (Word 2013 bug)
            if (WordDocument == null || !WordDocument.ActiveWindow.Visible || !WordDocument.ReadOnly || App.WordApplication.VersionLJ() <= OfficeVersion.Office2010)
                return;

            var view = WordDocument.ActiveWindow.View;
            var origViewType = view.Type;
            view.Type = WdViewType.wdNormalView;
            view.Type = WdViewType.wdPrintView;
            view.Type = origViewType;
        }


        public bool Editable { get; private set; }


        protected override void OnEditableChanged(EventArgs e)
        {
            base.OnEditableChanged(e);
            CheckDocumentStateChanged();
        }


        protected override void OnWindowViewTypeChanged(EventArgs e)
        {
            base.OnWindowViewTypeChanged(e);
            CheckDocumentStateChanged();
        }


        private void CheckDocumentStateChanged()
        {
            UserActionManager.UpdateDirty = true;

            //isValid = Me.WorkingWindowViewType = Word.WdViewType.wdPrintView OrElse Me.WorkingWindowViewType = Word.WdViewType.wdNormalView
            var isValid = IsEditable;
            if (Editable == isValid)
                return;

            if (!isValid)
            {
                //Do before setting _documentStateValid
                // Let the Transactions update the state first
                // in case they need to update other properties like the text or visible state.
                UserActionManager.Update();
            }

            Editable = isValid;

            UserActionManager.UpdateDirty = true;
            UserActionManager.Update();

            if (!Editable && Session.InSession)
            {
                ShowDocumentNotEditableMessage();
            }

            DocumentStateChanged?.Invoke(this, EventArgs.Empty);
        }


        protected virtual void ShowDocumentNotEditableMessage()
        {
            App.ViewService.ShowMessage("msgDocumentNotEditable");
        }


        public bool Dirty => ((IAddinAppDirty) this).DirtyCookie != 0;

        public virtual void MarkAsDirty() => MarkAsDirtyGenerator.MarkAsDirty(this);

        private int _dirtyCookie;

        int IAddinAppDirty.DirtyCookie
        {
            get => AutoDirty ? -1 : _dirtyCookie;
            set
            {
                if (!AutoDirty) _dirtyCookie = value;
            }
        }


        public AddinAppDocumentStorage Store { get; private set; }

        public virtual bool LicenseChecked { get; set; }

        protected virtual AddinAppDocumentStorage CreateStore() => new AddinAppDocumentStorage(this);

        /// <summary>
        /// Called when a Transaction is started and all initialization has been done. and fires the TransactionStarted event. It is not called for child transactions.
        /// </summary>
        /// <param name="e"></param>
        protected internal virtual void OnTransactionStarted(EventArgs e) => TransactionStarted?.Invoke(this, e);


        /// <summary>
        /// Called when a Transaction is Completed and fires the TransactionCompleted event. It is not called for child transactions.
        /// </summary>
        /// <param name="e"></param>
        protected internal virtual void OnTransactionCompleted(TransactionCompletedEventArgs e) => TransactionCompleted?.Invoke(this, e);


        protected internal bool OnUndoRequested()
        {
            return true;
            //if (App.ViewService.ShowMessage(App.Resources.GetString("CancelOpenTransactionMessage")) == DialogResult.Yes)
            //{
            //    return true;
            //}
        }


        protected internal void OnUndoCompleted()
        {
            UndoCompleted?.Invoke(this, EventArgs.Empty);
        }


#if (TRACK_DISPOSED)
/// <summary></summary>
        ~AddinAppDocument()
        {
#if (TRACK_DISPOSED)
            LevitJames.Core.Diagnostics.DisposeTracker.Add(_disposedSource);
#endif
            Dispose(disposing: false);
        }
#endif
 
    }
}
