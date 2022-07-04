// © Copyright 2018 Levit & James, Inc.

using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace LevitJames.MSOffice
{
    /// <summary>
    ///     Encapsulates the document properties of Word documents
    /// </summary>
    
    public sealed class DocumentPropertyLJ : IDisposable
    {
#if (TRACK_DISPOSED)
        private readonly string _disposedSource;
#endif


        /// <summary>
        ///     Creates a new DocumentPropertyLJ wrapping an existing Word document property
        /// </summary>
        /// <param name="prop">The Word document property to wrap</param>
        
        internal DocumentPropertyLJ(dynamic prop)
        {
            Instance = prop;
#if (TRACK_DISPOSED)
            _disposedSource = LevitJames.Core.Diagnostics.DisposeTracker.GetSource();
#endif
        }


        internal dynamic Instance { get; private set; }


        /// <summary>
        ///     Calls the "Delete" method of the wrapped property
        /// </summary>
        [DebuggerStepThrough]
        public void Delete()
        {
            Instance.Delete();
            Marshal.ReleaseComObject(Instance);
            Instance = null;
        }


        /// <summary>
        ///     Returns the "Creator" property of the wrapped property
        /// </summary>
        public int Creator => Convert.ToInt32(Instance.Creator);


        /// <summary>
        ///     Returns the Word.Application object associated with the wrapped property
        /// </summary>

        public Application Application => (Application) Instance.Application;


        /// <summary>
        ///     Returns the parent of the wrapped property
        /// </summary>
        public object Parent => Instance.Parent;


        /// <summary>
        ///     Provides access to the "Name" property of the wrapped property
        /// </summary>
        public string Name
        {
            get { return (string) Instance.Name; }
            set { Instance.Name = value; }
        }


        /// <summary>
        ///     Provides access to the "Value" property of the wrapped property
        /// </summary>
        public object Value
        {
            get { return Instance.Value; }
            set { Instance.Value = value; }
        }


        /// <summary>
        ///     Provides access to the "LinkToContent" property of the wrapped property
        /// </summary>
        
        
        
        public bool LinkToContent
        {
            get { return (bool) Instance.LinkToContent; }
            set { Instance.LinkToContent = value; }
        }


        /// <summary>
        ///     Provides access to the "LinkSource" property of the wrapped property
        /// </summary>
        
        
        
        public string LinkSource
        {
            get { return Instance.LinkSource; }
            set { Instance.LinkSource = value; }
        }


        /// <summary>
        ///     Provides access to the "Type" properly of the wrapped property
        /// </summary>
        
        
        
        [SuppressMessage("Microsoft.Naming", "CA1721:PropertyNamesShouldNotMatchGetMethods")]
        [CLSCompliant(isCompliant: false)]
        public MsoDocProperties Type
        {
            get { return (MsoDocProperties) Instance.Type; }
            set { Instance.Type = value; }
        }


        ~DocumentPropertyLJ()
        {
#if (TRACK_DISPOSED)
                LevitJames.Core.Diagnostics.DisposeTracker.Add(_disposedSource);
#endif
            Dispose(disposing: false);
        }

        private void Dispose(bool disposing)
        {
            if (Instance == null)
                return;
            // Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Marshal.ReleaseComObject(Instance);
            Instance = null;
            if (disposing)
                GC.SuppressFinalize(this);
        }

        /// <summary>
        ///     Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
        }
    }
}