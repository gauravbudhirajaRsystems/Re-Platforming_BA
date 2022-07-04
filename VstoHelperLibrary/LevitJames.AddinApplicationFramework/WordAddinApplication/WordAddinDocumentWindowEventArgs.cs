// © Copyright 2018 Levit & James, Inc.

using Microsoft.Office.Interop.Word;

namespace LevitJames.AddinApplicationFramework
{
    public class WordAddinDocumentWindowEventArgs : WordAddinDocumentEventArgs
    {
        public WordAddinDocumentWindowEventArgs(WordAddinDocument document, Window window) : base(document)
        {
            Window = window;
        }

        /// <summary>
        ///     Returns a Window instance
        /// </summary>



        public Window Window { get; }
    }
}
