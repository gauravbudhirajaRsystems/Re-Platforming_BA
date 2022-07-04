// © Copyright 2018 Levit & James, Inc.

using System.Diagnostics;

namespace LevitJames.AddinApplicationFramework
{
    /// <summary>
    ///     A base class that implements the IAddinAppProvider interface.
    ///     Inherited by classes that need a pointer to the AddinApplication/AddinDocument.
    ///     By inheriting from this interface the AddinApplication Storage will initialize the class after de-serialization or
    ///     creation with a pointer to the IAddinApplication.
    /// </summary>
    public abstract class AddinAppBase : IAddinAppProviderInternal
    {
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        internal IAddinApplicationInternal App { get; private set; }

        void IAddinAppProviderInternal.Initialize(IAddinApplicationInternal app, object data)
        {
            if (App != null)
                return;
            App = app;
            Initialize(data);
        }

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        IAddinApplication IAddinAppProvider.App => App;

        /// <summary>
        ///     Called after object de-serialization or on object creation by the AddinApplication framework, to set the
        ///     AddinAppDocument instance.
        /// </summary>
        /// <param name="data"></param>
        protected virtual void Initialize(object data) { }
    }
}