// © Copyright 2018 Levit & James, Inc.

using System.Collections.Generic;
using System.Diagnostics;
using JetBrains.Annotations;
using LevitJames.Core;
using Microsoft.Office.Interop.Word;

namespace LevitJames.MSOffice.MSWord
{
    /// <summary>
    ///     Class for maintaining a stack of <c>WordSettingsChangeSet</c> objects.
    /// </summary>
    
    public class WordSettingsChangeSetController
    {
        // FRONT MATTER
        private readonly Stack<WordSettingsChangeSet> _stack;
        private readonly Document _wordDoc;


        public WordSettingsChangeSetController(Document wordDoc)
        {
            _wordDoc = wordDoc;
            _stack = new Stack<WordSettingsChangeSet>();
        }


        /// <summary>
        ///     Method for creating a <c>WordSettingsChangeSet</c> object.
        /// </summary>
        /// <returns>Newly-created <c>WordSettingsChangeSet</c> object.</returns>
        /// <remarks>
        ///     <para>Once created, the <c>WordSettingsChangeSet</c> object is automatically pushed onto the stack.</para>
        ///     <seealso cref="Undo" />
        ///     <seealso cref="UndoAll" />
        /// </remarks>
        public WordSettingsChangeSet Create(string name = null)
        {
            if (string.IsNullOrEmpty(name))
            {
                var st = new StackTrace(fNeedFileInfo: false);
                name = st.GetFrame(index: 1).GetMethod().DeclaringType?.Name;
            }

            var changeSet = new WordSettingsChangeSet(_wordDoc, name);
            _stack.Push(changeSet);
            return changeSet;
        }


        /// <summary>
        ///     Method for restoring settings in a <c>WordSettingsChangeSet</c> object.
        /// </summary>
        /// <param name="changeSetToUndo"><c>WordSettingsChangeSet</c> object storing Word setting values to be restored.</param>
        /// <remarks>
        ///     <para>
        ///         <c>WordSettingsChangeSet</c> objects are popped off the stack until
        ///         the input <c>WordSettingsChangeSet</c> object is reached. As the objects
        ///         are popped off the stack, <c>UndoChangeSet</c> issues a <c>RestoreSettings</c>
        ///         call to the <c>WordSettingsChangeSet</c> object.
        ///     </para>
        ///     <seealso cref="Create" />
        ///     <seealso cref="UndoAll" />
        /// </remarks>
        public void Undo([NotNull] WordSettingsChangeSet changeSetToUndo)
        {
            Check.NotNull(changeSetToUndo, nameof(changeSetToUndo));

            if (_stack.Count == 0)
                return;

            Trace.TraceInformation("Undoing change sets in:" + changeSetToUndo.Name);

            WordSettingsChangeSet changeSet;
            do
            {
                changeSet = _stack.Pop();
                changeSet.Restore();
            } while (changeSet != changeSetToUndo);

            Trace.TraceInformation("Undoing change set out:" + changeSetToUndo.Name);
        }


        /// <summary>
        ///     Method for restoring all <c>WordSettingsChangeSet</c> objects in the stack.
        /// </summary>
        /// <remarks>
        ///     <para>
        ///         This method pops all <c>WordSettingsChangeSet</c> elements off the stack.
        ///         As elements are popped from the stack, a <c>RestoreSettings</c> call is
        ///         issued to restore the settings in the change set.
        ///     </para>
        ///     <seealso cref="Create" />
        ///     <seealso cref="Undo" />
        /// </remarks>
        public void UndoAll()
        {
            Trace.TraceInformation("UndoAll ChangeSets in");

            if (_stack.Count > 0)
            {
                do
                {
                    var changeSet = _stack.Pop();
                    changeSet.Restore();
                } while (_stack.Count != 0);
            }

            Trace.TraceInformation("UndoAll ChangeSets out");
        }
    }
}