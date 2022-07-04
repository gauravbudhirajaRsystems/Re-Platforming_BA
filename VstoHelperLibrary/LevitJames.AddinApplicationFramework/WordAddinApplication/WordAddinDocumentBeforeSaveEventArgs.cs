// © Copyright 2018 Levit & James, Inc.

namespace LevitJames.AddinApplicationFramework
{
    public class WordAddinDocumentBeforeSaveEventArgs : WordAddinDocumentEventArgs
    {
        public WordAddinDocumentBeforeSaveEventArgs(WordAddinDocument document, bool saveAsUI, bool cancel)
            : base(document)
        {
            SaveAsUI = saveAsUI;
            Cancel = cancel;
        }

        public bool Cancel { get; set; }
        public bool SaveAsUI { get; set; }
    }
}
