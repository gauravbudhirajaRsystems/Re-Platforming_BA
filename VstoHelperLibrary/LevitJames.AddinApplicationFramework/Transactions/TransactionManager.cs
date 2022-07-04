//**************************************************
//** © 2021 Litera Corp. All Rights Reserved.
//***************************************************

using System;
using System.Diagnostics;
using System.Runtime.Caching;
using LevitJames.AddinApplicationFramework.Properties;
using LevitJames.Core;
using LevitJames.MSOffice.MSWord;

namespace LevitJames.AddinApplicationFramework
{
    internal sealed class TransactionManager : AddinAppBase, IUserActionTarget
    {
        private readonly TransactionFactory _factory;
        private bool _appValidationDone;


        private MemoryCache _cache;

        private bool _inEndSession;
        private TransactionMetadata _queuedMetadata;


        public TransactionManager()
        {
            _factory = new TransactionFactory();
        }

        public MemoryCache Cache => _cache ?? (_cache = new MemoryCache("StorageManager"));


        public AddinAppTransactionBase ActiveTransaction => App?.ActiveDocument?.Session?.ActiveTransaction;


        void IUserActionTarget.UserActionExecuted(UserActionExecuteEventArgs e)
        {
            if (e.Handled)
                return;

            try
            {
                if (e.Id == AddinAppUserActionConstants.CloseSession && App.InSession)
                {
                    e.Handled = true;
                    App.ActiveDocument.Session.CloseNamedSession();
                    return;
                }

                var result = ExecuteTransaction(new TransactionMetadata(e.UserAction ?? (object)e.Id, e.Parameter));

                if (result)
                    WordExtensions.BeginInvoke<AddinAppDocument>(ExecuteQueuedTransaction, App.ActiveDocument);
                e.Handled = result;
            }
            catch (Exception exception)
            {
                e.Handled = true; // make sure we set to true so exception is not thrown by UserActionManager
                App.OnUnhandledException(exception);
            }

        }

        object IUserActionTarget.Context => App;


        //initialization routines
        protected override void Initialize(object data)
        {
            base.Initialize(data);
            ((IAddinAppProviderInternal)_factory).Initialize(App);

            AddinAppSession.AddCloseSessionUserActions();

            UserActionManager.AddTarget(this);
            UserActionManager.ActiveTarget = this;
        }


        //start/end Session


        private bool AppValidation()
        {
            if (_appValidationDone)
                return true;

            // Validate Word version
            if (!App.Environment.IsWordCompatible)
            {
                var msg = App.GetStringResource(nameof(Resources.aafmsgWordVersionIncompatible));
                App.ViewService.ShowMessage(msg);
                return false;
            }

            // Check for valid Word.Application.Username
            if (!App.CheckWordUserNameIsValid())
                return false;

            _appValidationDone = true;
            return true;
        }


        /// <summary>
        ///     Ends the session started using the StartSession Call
        /// </summary>

        public void EndEditSession(AddinAppDocument document)
        {
            if (_inEndSession)
                return;

            if (document == null || !document.Session.InSession)
                return;

            App.Tracer.TraceInformation("EndEditSession in");

            var doc = document.WordDocument;
            var trackRevisions = doc?.TrackRevisionsLJ(value: false);
            App.Tracer.TraceInformation("Revisions disabled");
            var transaction = ActiveTransaction;
            try
            {
                _inEndSession = true;

                if (transaction?.Result == TransactionResult.None)
                {
                    App.Tracer.TraceInformation("Abandoning transaction");
                    transaction.AbandonTransaction();
                }

                if (!document.Session.InNamedSession())
                {
                    App.Tracer.TraceInformation("Exiting all sessions");
                    document.ExitAllSessions();
                }
                else
                {
                    // ReSharper disable once RedundantAssignment
                    var currentSessionName = document.Session?.Current?.Name ?? "(None)";
                    App.Tracer.TraceInformation($"In named session: {currentSessionName}");
                }
            }
            finally
            {
                App.ViewService.Progress.Close();
                _inEndSession = false;

                if (trackRevisions != null)
                {
                    App.Tracer.TraceInformation($"Resetting TrackRevisions: {trackRevisions.Value}");
                    doc.TrackRevisionsLJ(trackRevisions.Value);
                }
            }

            App.Tracer.TraceInformation("EndEditSession out");
        }


        private void CreateRecoveryLog(AddinAppDocument document)
        {
            if (!App.UserSettings.Special.CreateWordRecoveryLog)
                return;

            var recoveryLog = new WordObjectModelRecovery();
            recoveryLog.Save(document.WordDocument, RecoveryScope.All, App.Paths.WordRecoveryLogFile);
        }


        internal bool ExecuteTransaction(TransactionMetadata metadata)
        {
            ActiveTransaction?.AbandonTransaction();

            var document = App.ActiveDocument;
            TransactionMetadata resolvedMetadata = null;

            var transaction = _factory.Create(metadata);
            if (transaction == null)
            {
                App.AppTracer.Tracer.TraceEvent(TraceEventType.Error, 0, "Cannot create Transaction for Id :" + metadata.Id);
                return false;
            }

            if (!transaction.IsLightweight())
            {
                if (document != null && !document.Session.InSession)
                {
                    //Do AppValidation before other checks as we don't want to be doing licensing stuff
                    //if the Word Version is not compatible etc.
                    // App Level validation on entering an edit session.
                    if (!AppValidation())
                        return true;
                }

                if (document != null && !document.LicenseChecked)
                {
                    var licType = _factory.TransactionTypeFromUserAction(AddinAppUserActionConstants.Licensing);
                    if (licType != null && licType.IsSubclassOf(typeof(LicensingTransaction)))
                        resolvedMetadata = new TransactionMetadata(AddinAppUserActionConstants.Licensing, null, metadata);
                }


                //NOTE: Any transactions resolved above, to run *before* the InitialEditTransaction must be marked with the 
                //      LightWeightTransactionAttribute otherwise we will enter a Session without calling the
                //      InitialEditSession code below.
                //TODO: NJKA: Add an Exception/Debug.Assert to check this.
                if (resolvedMetadata == null && document != null && !document.Session.InSession)
                {
                    //If this is the first edit session fire the InitialEditTransaction transaction
                    if (!document.InitialEnterSessionDone)
                    {
                        CreateRecoveryLog(document);
                        //Need to change the meta data to queue the InitialEditSession transaction first;
                        //Set root transaction to passed in meta data;
                        resolvedMetadata = new TransactionMetadata(AddinAppUserActionConstants.InitialEditSession, null, metadata);
                    }
                }
            }


            resolvedMetadata = resolvedMetadata ?? metadata;
            if (resolvedMetadata != transaction.Metadata)
                transaction = _factory.Create(resolvedMetadata); // create regular


            ExecuteTransaction(transaction);
            return true;
        }


        private void ExecuteTransaction(AddinAppTransactionBase transaction)
        {
            if (transaction == null)
                return;

            App.OnTransactionStarting(transaction.Metadata);
            App.Tracer.TraceInformation("ExecuteTransaction start: " +
                                        (transaction.Name));

            var document = transaction.Document;
            var isLightweight = transaction.IsLightweight();
            var exception = false;

            if (document != null)
                document.Session.ActiveTransaction = transaction;

            if (!isLightweight && document != null && !document.Session.InSession)
            {
                document.EnterSession();
            }


            document?.SetExclusiveEventsFlag(add: true);

            UserActionManager.Enabled = false;
            TransactionResult result;
            try
            {
                var session = transaction.NamedSession;
                transaction.TransactionCompleted += TransactionCompletedHandler;
                transaction.Initiate();
                result = transaction.Result;

                //Transaction completed at this point or it is modeless
                TryEndEditSession(transaction);

                if (transaction.Document != null && result == TransactionResult.Success)
                {
                    if (session != null && transaction.Document.Session.InNamedSession(session))
                    {
                        //TODO - NJKA: Is this needed now?
                        transaction.ResetTransactionResult();
                    }
                }
            }
            //UserCanceledOperationException is caught and handled inside transaction.Initiate()
            catch (Exception ex)
            {
                _queuedMetadata = null;
                exception = true;

                App.Tracer.TraceInformation("Exception in ExecuteTransaction: " + ex.Message);

                // Restore Payne's Numbering Assistant
                //WordHelper.SetPayneNumberingAssistantState(DocData, enabled: true);
                document?.Store.OnCompleteTransaction();

                UserActionManager.UpdateDirty = true;

                throw;
            }
            finally
            {
                UserActionManager.Enabled = true;
                UserActionManager.UpdateIfDirty();

                //Don't call OnClearModel if InSession is true, as it's too complex to reset the model.
                //Alternative might be to call document.Session.Close()
                if (document != null && document.Session?.InSession == false && (_queuedMetadata == null || exception))
                    document.OnClearModel();

                //Document may have been closed
                if (document?.WordDocument != null && App.WordApplication.IsObjectValid(document.WordDocument))
                    document.SetExclusiveEventsFlag(add: false);

                //Keep last as RefreshRibbon
                if (App.ViewService.RibbonInvalidated)
                    App.ViewService.RefreshRibbon();

                if (_queuedMetadata == null)
                    App.ViewService.Progress.Close();
            }

            if (result != TransactionResult.None) // modeless
                App.OnTransactionCompleted(transaction.Metadata, _queuedMetadata != null);


            App.Tracer.TraceInformation("ExecuteTransaction end: " + transaction.Name);
        }


        private void TryEndEditSession(AddinAppTransactionBase transaction)
        {
            if (transaction.Document == null || (transaction.InOnInitiate || transaction.Result == TransactionResult.None))
                return; // If transaction.InOnInitiate then this method will be called later.

            if (transaction.Document.Store.ClearAllCalled)
            {
                //Always EndEditSession if ClearAllCalled = true;
                transaction.Document.Store.ClearAllCalled = false;
            }
            else if (_queuedMetadata != null)
            {
                return;
            }

            EndEditSession(transaction.Document);
        }

        private void SetNextTransactionMetadata(AddinAppTransactionBase transaction, TransactionCompletedEventArgs e)
        {
            if (transaction.Result != TransactionResult.Success || string.IsNullOrEmpty(e.NextUserActionId))
            {
                //This is the last transaction so clear the meta data
                _queuedMetadata = null;
                return;
            }

            //Fill a new TransactionMetadata with the new and current Transaction information.
            _queuedMetadata = new TransactionMetadata(e.NextUserActionId, e.NextUserActionParameter, transaction.Metadata);
        }


        private void TransactionCompletedHandler(object sender, TransactionCompletedEventArgs e)
        {
            var transaction = (AddinAppTransactionBase)sender;
            var namedSession = transaction.NamedSession;

            var document = transaction.Document;

            transaction.TransactionCompleted -= TransactionCompletedHandler;

            SetNextTransactionMetadata(transaction, e);

            if (document != null)
            {
                var documentSession = document.Session;
                if (namedSession != null)
                {
                    if (transaction.Result != TransactionResult.Success && documentSession.InNamedSession(namedSession))
                        //Close the current named session if it belongs to Active transaction and transaction was cancelled.
                        //AKA Transaction.PushTransaction was not called put the transaction then failed or was cancelled.
                        documentSession.CloseNamedSession();

                    //Don't automatically call PushNamedSession.
                    // Let the dev explicitly push the named session.

                }

                documentSession.ActiveTransaction = null;
            }

            TryEndEditSession(transaction);

            document?.OnTransactionCompleted(e);

            App.ViewService.InvalidateRibbon();
            UserActionManager.UpdateDirty = true;

            if (_queuedMetadata == null)
            {

                App.ViewService.Progress.Close();
            }


            if (_queuedMetadata != null && !transaction.InOnInitiate)
            {
                //We are being called outside of the InOnInitiate call so we need to execute the
                // queued transaction from here.

                //Force refresh of the ribbon.
                App.ViewService.RefreshRibbon();

                App.OnTransactionCompleted(transaction.Metadata, hasQueuedTransaction: true);

                WordExtensions.BeginInvoke<AddinAppDocument>(ExecuteQueuedTransaction, App.ActiveDocument);
            }

        }

        private void ExecuteQueuedTransaction(AddinAppDocument activeDocument)
        {
            try
            {
                App.ViewService.InvalidateRibbon();
                App.ViewService.RefreshRibbon();

                if (_queuedMetadata == null)
                    return;

                if (!activeDocument.IsActive)
                    activeDocument.WordDocument.ActivateWindowLJ();

                var metaData = _queuedMetadata;
                _queuedMetadata = null;

                ExecuteTransaction(metaData);

                if (_queuedMetadata != null)
                    WordExtensions.BeginInvoke<AddinAppDocument>(ExecuteQueuedTransaction, App.ActiveDocument);
            }
            catch (Exception e)
            {
                //BA-1409
                //Because this method can be called from BeginInvoke we need to catch the exception here.
                App.OnUnhandledException(e);
            }
        }
 
        public AddinAppTransactionBase CreateTransaction(TransactionMetadata metadata) => _factory.Create(metadata);


        public bool RegisterTransaction(object userAction, Type transactionType) => _factory.RegisterTransaction(userAction, transactionType);
    }
}
