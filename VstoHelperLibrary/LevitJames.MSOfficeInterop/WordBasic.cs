// © Copyright 2018 Levit & James, Inc.

using Microsoft.Office.Interop.Word;

namespace LevitJames.MSOffice.MSWord
{
    /// <summary>
    ///     Provides methods from the old WordBasic model that are not available in the new Word object model.
    /// </summary>
    public sealed class WordBasic
    {
        internal WordBasic(Application source)
        {
            Instance = source.WordBasic;
        }

        internal dynamic Instance { get; }


        /// <summary>
        ///     Gets the file name of a Document even if it was opened under source control.
        /// </summary>
        public string FileName => Instance.Filename;


        /// <summary>
        ///     Disables Word macros from running when a document is opened or a new document is created.
        /// </summary>
        public void DisableAutoMacros()
        {
            if (Instance == null)
                return;

            Instance.DisableAutoMacros(1);
        }

        /// <summary>
        ///     Enables Word macros from running when a document is opened or a new document is created.
        /// </summary>
        public void EnableAutoMacros()
        {
            if (Instance == null)
                return;

            Instance.DisableAutoMacros(0);
        }
    }
}
