using System;
// © Copyright 2018 Levit & James, Inc.

namespace LevitJames.AddinApplicationFramework
{
    public class WordAddinDocumentEventArgs : EventArgs
    {
        public WordAddinDocumentEventArgs(WordAddinDocument document)
        {
            Document = document;
        }

        public WordAddinDocument Document { get; }
    }
}
