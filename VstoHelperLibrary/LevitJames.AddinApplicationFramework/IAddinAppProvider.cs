// © Copyright 2018 Levit & James, Inc.

namespace LevitJames.AddinApplicationFramework
{
    public interface IAddinAppProvider
    {
        IAddinApplication App { get; }
    }
 
    internal interface IAddinAppProviderInternal : IAddinAppProvider
    {
        void Initialize(IAddinApplicationInternal app, object data = null);
    }
}