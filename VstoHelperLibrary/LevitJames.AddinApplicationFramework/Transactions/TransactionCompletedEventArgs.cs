// © Copyright 2018 Levit & James, Inc.

using System;

namespace LevitJames.AddinApplicationFramework
{
    public class TransactionCompletedEventArgs : EventArgs
    {
        internal TransactionCompletedEventArgs(AddinAppTransactionBase transaction, TransactionResult result, string action, object parameter)
        {
            Transaction = transaction;
            Result = result;
            NextUserActionId = action;
            NextUserActionParameter = parameter;
        }

        public string Id => Transaction.Metadata.Id;

        public AddinAppTransactionBase Transaction { get; }

        public string NextUserActionId { get; }
        public object NextUserActionParameter { get; }

        // ReSharper disable once ConvertToVbAutoProperty
        public TransactionResult Result { get; }
    }
}