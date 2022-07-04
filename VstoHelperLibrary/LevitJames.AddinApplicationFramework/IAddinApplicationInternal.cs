// © Copyright 2018 Levit & James, Inc.

using System.Collections.Generic;
using System.Reflection;
using LevitJames.Core;

namespace LevitJames.AddinApplicationFramework
{
    internal interface IAddinApplicationInternal : IAddinApplication
    {

        TransactionManager TransactionManager { get; }

        AddinAppTracer AppTracer { get; }
        IEnumerable<Assembly> GetTransactionAssemblies();
        void ConfigureTracing();

        void OnStorageItemNotVerified(AddinAppTransactionBase transaction, string key, string hint, IEnumerable<SerializationStreamComparerSet> items, out bool ignore);

    }
}