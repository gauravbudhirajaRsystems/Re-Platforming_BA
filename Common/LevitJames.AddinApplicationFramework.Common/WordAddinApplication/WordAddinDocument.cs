using LevitJames.MSOffice.Common;
using System;

namespace LevitJames.AddinApplicationFramework.Common.WordAddinApplication
{
    public abstract class WordAddinDocument : IWordDocumentProvider, IDisposable
    {
        public void Dispose()
        {
            throw new NotImplementedException();
        }
    }
}
