//*************************************************
//* © 2021 Litera Corp. All Rights Reserved.
//**************************************************

using System;
using System.Diagnostics;
using System.Resources;
using LevitJames.Core;
using Microsoft.Office.Interop.Word;
using Version = System.Version;

namespace LevitJames.AddinApplicationFramework
{
    public interface IAddinApplication
    {

        event EventHandler Idle;
        event EventHandler Closing;

        bool InSession { get; }

        Application WordApplication { get; }

        Version Version { get; }

        ResourceManager Resources { get; }

        bool HasActiveDocument { get; }

        AddinAppDocument ActiveDocument { get; }

        TraceSource Tracer { get; }
        AddinAppEnvironment Environment { get; }
        AddinAppPaths Paths { get; }
        AddinAppUserSettings UserSettings { get; }
        AddinAppAdministrativeSettings AdminSettings { get; }

        AddinAppViewService ViewService { get; }

        event EventHandler StartupCompleted;
        void RegisterTransaction(object id, Type transaction);

        void EnsureActiveDocument();

        bool CheckWordUserNameIsValid();

        void OnUnhandledException(Exception ex);

        string GetStringResource(string resourceName, params string[] args);
        AddinAppDocumentRecovery CreateDocumentRecovery();

        void OnTransactionStarting(TransactionMetadata metadata);
        void OnTransactionCompleted(TransactionMetadata metadata, bool hasQueuedTransaction);

        void Initialize(AddinAppBase instance, object data = null);
    }
}