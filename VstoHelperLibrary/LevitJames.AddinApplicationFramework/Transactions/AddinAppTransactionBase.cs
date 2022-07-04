//*************************************************
//* © 2021 Litera Corp. All Rights Reserved.
//**************************************************

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using LevitJames.Core;
using LevitJames.MSOffice.MSWord;
using LevitJames.TextServices;

namespace LevitJames.AddinApplicationFramework
{
    /// <summary>
    ///     Creates the required Transaction class that is Registered to the UserAction else returns nothing.
    /// </summary>
    /// <remarks>
    ///     In order for a UserAction to automatically create a Transaction the Transaction must obey the following rules.
    ///     The userAction passed must have been registered within the transaction from the Shared method that uses the
    ///     CreateUserActionsAttribute.
    ///     The Transaction must contain one or more of the following constructors. The following list is the order in which a
    ///     valid constructor is found. The constructor can be Public or Friend or Private.
    ///     <list>
    ///         <listheader>Valid Constructors are</listheader>
    ///         <pageItem> userAction , parameter </pageItem>
    ///         <pageItem>userAction</pageItem>
    ///         <pageItem>userActionId , parameter</pageItem>
    ///         <pageItem>userActionId</pageItem>
    ///         <pageItem>parameter</pageItem>
    ///         <pageItem>(none)</pageItem>
    ///     </list>
    /// </remarks>
    [Serializable]
    public abstract class AddinAppTransactionBase : AddinAppBase
    {
        private List<Tuple<AddinAppHistoryRecord, bool>> _history;
        private bool _inNamedSession;
 
        private WordState _state;

        protected IAddinApplication AddinApp => App;



        public TransactionMetadata Metadata { get; private set; }


        /// <summary>
        ///     Provides a friendly name for the Transaction. The default implementation takes the Camel cased Transaction name
        ///     inserts spaces into it.
        ///     Override to provide an alternative descriptive name.
        /// </summary>
        
        
        
        public virtual string Name => GetType().Name.ToSpacedCase(trimStart: null, trimEnd: "Transaction", ignoreCase: true,
                                                                  style: SpacedCaseStyle.CapitalizeAllWords);


        public AddinAppDocument Document { get; private set; }


        internal bool InOnInitiate { get; set; }


        protected AddinAppTransactionBase ChildTransaction { get; private set; }


        internal bool IsNestedTransaction => Metadata?.Parent != null;


        /// <summary>
        ///     You must override this in each transaction class, so that it can be queried.  It indicates how the transaction
        ///     interacts with the Undo/Redo
        ///     stack (e.g. is it Undoable?  Is it Informational only, etc.)
        /// </summary>
        public abstract TransactionType Type { get; }


        public TransactionResult Result { get; private set; }


        protected internal bool Undoable => Document != null && (Type == TransactionType.ReadWriteUndoable || Type == TransactionType.ReadWrite) && _inNamedSession;


        protected internal virtual bool RestoreStateOnCancel => false;
  
        public event EventHandler<TransactionCompletedEventArgs> TransactionCompleted;

        protected override void Initialize(object data)
        {
            base.Initialize(data);
            Document = App.ActiveDocument;
            Metadata = (TransactionMetadata) data;
        }

        /// <summary>
        ///     Returns an optional IAddinAppNamedSession instance. Used when the Transaction is starting a named session. For
        ///     example ReviewMode.
        /// </summary>
        /// <remarks>
        ///     The first this this property is called it will call CreateNamedSession(). Override the CreateNamedSession to create
        ///     the class that implements the IAddinAppNamedSession interface.
        /// </remarks>
        protected internal IAddinAppNamedSession NamedSession { get; private set; }

        /// <summary>
        ///     Pushes the named session for this transaction, making it the Current Named session.
        /// </summary>
        /// <remarks>
        ///     This can be called anywhere in the OnInitiate override when code relies on the current named session to be active.
        ///     For example when a named session transaction needs to call child transactions.
        ///     <para>
        ///         If this method is not called during the Initiate call then it will be called as required when the
        ///         CompleteTransaction method is called.
        ///     </para>
        /// </remarks>
        protected internal void PushNamedSession(IAddinAppNamedSession namedSession)
        {
            if (NamedSession != null)
                throw new InvalidOperationException("NamedSession already set");
            NamedSession = namedSession;
            Document.Session.PushNamedSession(namedSession);
        }

        //protected internal virtual IAddinAppNamedSession CreateNamedSession() => null;

        protected void SetWordState(WordStateOptions states)
        {
            if (_state != null)
                throw new InvalidOperationException("Already set");

            _state = new WordState(Document.WordDocument, states, delay: false);
        }

        protected void RestoreWordState()
        {
            if (_state != null)
            {
                _state.Restore();
                _state = null;
            }
        }


        //This routine should be kept none overridable as it controls the specific order that transactions are Initialized. 
        // Overridable code should be place in additional overridable methods or placed in the OnInitiate method.
        /// <summary>
        ///     Called to Initiate (run) the Transaction.
        /// </summary>
        /// <remarks>This method does all the compulsory initialization before calling the overridable OnInitiate member.</remarks>
        protected internal void Initiate()
        {
            if (Type == TransactionType.Unknown)
                throw new InvalidOperationException("TransactionType not set");

            App.Tracer.TraceInformation("TransactionBase:Initiate:" + GetType().Name);
            Trace.IndentSize += 1;
            var priorFocus = App.ViewService.GetFocus();

            try
            {
                InOnInitiate = true;
                _inNamedSession = Document?.Session?.InNamedSession() == true;
                Document?.Store?.BeginTransaction();
                App.Tracer.TraceInformation("TransactionBase:OnInitiate: in" + Name);
                if (!IsNestedTransaction)
                    Document?.OnTransactionStarted(EventArgs.Empty);

                if (Document != null && App.WordApplication.ActiveDocumentLJ() == null)
                    Document.WordDocument.ActivateWindowLJ();

                OnInitiate();
                App.Tracer.TraceInformation("TransactionBase:OnInitiate: out" + Name);
            }
            catch (Exception ex)
            {
                InOnInitiate = false;

                //Also check GetBaseException in-case exception is wrapped and then re-thrown.
                if (ex is UserCanceledOperationException || ex.GetBaseException() is UserCanceledOperationException)
                {
                    App.Tracer.TraceInformation("User Canceled from progress dialog: ");
                    if (Result == TransactionResult.None)
                    {
                        CompleteTransaction(TransactionResult.Cancel);
                    }

                    //Handled so we do not throw
                }
                else
                {
                    throw;
                }
            }
            finally
            {
                InOnInitiate = false;
                //Needs to be placed after the OnInitiate as modal dialogs must have returned from the ModalLoop before we can set this here
                if (Result != TransactionResult.None)
                {
                    OnRestoreFocus(priorFocus);
                }

                Trace.IndentSize -= 1;
            }

            App.Tracer.TraceInformation("Base Initiate end: " + Name);
        }


        /// <summary>
        ///     Restores the focus that was saved at the start of the Transaction.
        /// </summary>
        /// <param name="priorFocus">The focus that was captured</param>
        /// <remarks>Override if the prior focus should not be restored, or different behaviour is required.</remarks>
        protected virtual void OnRestoreFocus(object priorFocus) => App.ViewService.SetFocus(priorFocus);


        /// <summary>
        ///     Called from Initiate after the base implementation has done all the required initialization.
        /// </summary>
        
        protected abstract void OnInitiate();


        //Child Transactions


        protected TransactionResult InitiateChildTransaction(string userActionId, object parameter = null)
        {
            Check.NotEmpty(userActionId, "userActionId");
            var childMetadata = new TransactionMetadata(userActionId, parameter, this);
            var transaction = App.TransactionManager.CreateTransaction(childMetadata);
            return InitiateChildTransaction(transaction);
        }
 
        /// <summary>
        ///     Called to Initiate a child, nested transaction, so that the child transaction can rolled back if it fails or is
        ///     canceled.
        /// </summary>
        /// <param name="transaction">The child Transaction to Initiate</param>

        private TransactionResult InitiateChildTransaction(AddinAppTransactionBase transaction)
        {
            Check.NotNull(transaction, nameof(transaction));
            Check.NotNull(transaction.Metadata, nameof(transaction.Metadata));

            if (transaction == this)
                throw new InvalidOperationException("Invalid Transaction passed.");

            if (ChildTransaction != null)
                throw new InvalidOperationException("Only one child transaction allowed.");

            if (transaction.NamedSession != null)
                throw new InvalidOperationException("Child Transactions cannot be named session transactions.");
 
            if (transaction.Metadata.Parent != this)
                throw new InvalidOperationException("Invalid Parent.");

            App.Tracer.TraceInformation("InitializeChildTransactionCore start: " + transaction.Name);

            ChildTransaction = transaction;
            ChildTransaction.TransactionCompleted += (sender, args) =>
                                                     {
                                                         ChildTransaction = null;
                                                         if (Result == TransactionResult.None && Document?.Session != null) {
                                                             Document.Session.ActiveTransaction = this;
                                                         }
                                                             
                                                     };

            var exceptionWasThrown = false;
            if (Document?.Session != null)
                Document.Session.ActiveTransaction = transaction;
            try
            {
                UserActionManager.UpdateDirty = true;
                transaction.Initiate();
            }
            catch (Exception ex)
            {
                exceptionWasThrown = true;

                if (ex.GetBaseException() is UserCanceledOperationException)
                    throw new UserCanceledOperationException();

                throw;
            }
            finally
            {
                if (exceptionWasThrown || transaction.Result != TransactionResult.None)
                {
                    ChildTransaction = null;
                    if (Document?.Session != null)
                        Document.Session.ActiveTransaction = this;
                }
            }

            App.Tracer.TraceInformation("InitializeChildTransactionCore completed: " + transaction.Name);

            return transaction.Result;
        }

        internal AddinAppTransactionBase NestedTransactionOrThis()
        {
            var nested = this;
            while (true)
            {
                if (nested?.ChildTransaction == null)
                    return nested;

                nested = nested.ChildTransaction;
            }
        }


        /// <summary>
        ///     Called to abandon this transaction
        /// </summary>
        
        protected internal virtual void AbandonTransaction()
        {
            //If Result is already Abandoned then we already tried to call CompleteTransaction(TransactionResult.Abandoned)
            if (Result == TransactionResult.Abandoned)
            {
                return;
            }

            CompleteTransaction(TransactionResult.Abandoned);
        }


        /// <summary>
        ///     Returns if this transaction is lightweight. Meaning a transaction that does not require storage, undo or rollback
        ///     support.
        /// </summary>
        
        public bool IsLightweight() => (Type == TransactionType.ReadOnly) && Metadata.Id != AddinAppUserActionConstants.Undo;

 
        internal void ResetTransactionResult()
        {
            Result = TransactionResult.None;
        }

#if !IGNORE_DEBUGGER_HIDDEN_ATTRIBUTE
        [DebuggerHidden]
#endif

        protected bool? ShowView(IViewModel model, object owner = null) => App.ViewService.Show(model, owner);


        /// <summary>
        ///     Completes the transaction and raises the TransactionCompleted event.
        /// </summary>
        /// <param name="result">The result of the transaction. Cannot be TransactionResult.None.</param>
        /// <param name="nextUserAction">
        ///     The name of a UserAction to execute after completion. The nextUserAction is only used if
        ///     the result passed equals TransactionResult.Success.
        /// </param>
        /// <param name="nextUserActionParameter">A parameter to pass to the next UserAction defined in nextUserAction</param>
        
        protected virtual void CompleteTransaction(TransactionResult result, string nextUserAction = null, object nextUserActionParameter = null)
        {
            if (result == TransactionResult.None)
                throw new ArgumentException(@"result cannot be none.", nameof(result));

            if (Result == TransactionResult.Abandoned)
                return; // likely from an event handler.

            if (Result != TransactionResult.None)
                throw new InvalidOperationException("CompleteTransaction has already been called");

            App.Tracer.TraceInformation("CompleteTransaction (base) in: " + nextUserAction);

            if (ChildTransaction != null)
            {
                ChildTransaction.CompleteTransaction(result, nextUserAction, nextUserActionParameter);
                return;
            }

            //Set the final Transaction result.
            Result = result;

            App.ViewService.OnOwnerClosing(this);

            OnInsertHistoryRecords();

            //if (Result != TransactionResult.Success)
            //{
            //    nextUserAction = null;
            //    nextUserActionParameter = null;
            //}

            if (Undoable)
                Document.App.ViewService.Progress.Reporter?.Increment(message: "Saving...");

            Document?.Store?.OnCompleteTransaction();

            OnTransactionCompleted(new TransactionCompletedEventArgs(this, Result, nextUserAction, nextUserActionParameter));

            RestoreWordState();

            App.Tracer.TraceInformation("CompleteTransaction (base) done");
        }

        protected void CompleteTransaction(bool? succeeded, string successMessage = null, string nextTransaction = null, object nextTransactionParameter = null)
        {
            CompleteTransaction(succeeded == true ? TransactionResult.Success : TransactionResult.Cancel, nextTransaction, nextTransactionParameter);
            if (succeeded == true && !string.IsNullOrEmpty(successMessage))
            {
                App.ViewService.ShowMessage(successMessage);
            }
        }


        /// <summary>
        ///     Adds the supplied record text and item text describing the transaction into the Document.History collection.
        /// </summary>
        /// <param name="details">The name of the record to add</param>
        /// <param name="text"></param>
        /// <param name="addEvenIfTransactionResultNotSuccess">Writes history record even if Transaction.Result is not Success.</param>
        protected void AddHistoryRecord(string text, string details = null, bool addEvenIfTransactionResultNotSuccess = false)
        {
            Debug.Assert(!string.IsNullOrEmpty(text), "Empty history record text");

            // History Levels:
            // 0	- History record for the entire App session
            // 1	- History record for transactions not in a named session
            // 2+	- History record for transactions in a named session
            var historyLevel = Document.Session.NamedSessionCount() + 1;
            var record = new AddinAppHistoryRecord(historyLevel, text, details);

            _history = _history ?? new List<Tuple<AddinAppHistoryRecord, bool>>();
            _history.Add(new Tuple<AddinAppHistoryRecord, bool>(record, addEvenIfTransactionResultNotSuccess));
        }


        /// <summary>
        ///     Writes out the a record of the transaction to the Documents HistoryRecord collection.
        ///     The default name for the history name is taken from the Transaction.Name property.
        ///     For transactions with multiple actions the method can be overridden and a a more specific name can be supplied
        ///     along with item text if required.
        /// </summary>
        protected virtual void OnInsertHistoryRecords()
        {
            // Don't write a blank history record, default to the transaction name.
            //For transactions with multiple user actions override and supply a more specific name
            if (_history == null && Result != TransactionResult.Success)
                return;

            if (_history == null)
            {
                var autoWrite = Type == TransactionType.ReadWrite || Type == TransactionType.ReadWriteUndoable || this is UndoTransaction;
                if (autoWrite)
                    AddHistoryRecord(Name);
            }

            if (_history == null)
                return;

            var itemsToAdd = Result == TransactionResult.Success ? _history : _history.Where(i => i.Item2);

            //Only add the items marked addEvenIfTransactionResultNotSuccess.
            foreach (var itm in itemsToAdd)
                Document.History.AddRecord(itm.Item1);

            _history.Clear();
        }


        /// <summary>
        ///     Raises the TransactionCompleted event.
        /// </summary>
        /// <param name="e"></param>
        
        protected virtual void OnTransactionCompleted(TransactionCompletedEventArgs e)
        {
            TransactionCompleted?.Invoke(this, e);
        }

        protected internal virtual void OnUndoCompleted() { }
    }
}