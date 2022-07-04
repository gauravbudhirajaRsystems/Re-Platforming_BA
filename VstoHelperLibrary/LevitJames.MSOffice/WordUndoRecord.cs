// © Copyright 2018 Levit & James, Inc.

using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using LevitJames.Core;
using Microsoft.Office.Interop.Word;

//© Copyright 2009 Levit & James, Inc.

namespace LevitJames.MSOffice.MSWord
{
    /// <summary>
    ///     A class used to provide extended undo functionality to Word.
    /// </summary>
    /// <remarks>
    ///     Use StartCustomRecord,EndCustomRecord, SuspendRecording/ResumeRecording to chain Word UndoRecords, which would
    ///     otherwise be ended by certain Word Object model calls or by DocumentChange events.
    ///     StartCustomRecord/EndCustomRecord etc. are only applicable to word versions 14 (2010) and greater. However calls
    ///     are safe to call in earlier versions of Word.
    ///     <para>
    ///         Use StartUndoMarker/EndUndoMarker in any versions of Word to add an UndoMarker. Calling EndUndoMarker will undo
    ///         all the Word Document actions back to the point where StartUndoMarker was called.
    ///     </para>
    /// </remarks>
    public class WordUndoRecord
    {
        private const string UndoStyleName = "_LJ_WUR_";
        private Document _document;
        private int _level;
        private int _partNumber;

        private LockCounter _suspended;
        private LockCounter _undoMarkerLevel;

        // 1 means in StartUndoMarker
        // 2 means in StartUndoMarker and recording was suspended so resume when EndUndoMarker is called.
        private int _undoMarkerState;


        internal WordUndoRecord() { }


        public int Id { get; private set; }


        /// <summary>
        ///     Returns True if SuspendRecording has been called causing Word UndoRecord recording to be suspended.
        /// </summary>
        
        
        
        public bool IsSuspended => _suspended.Locked;


        /// <summary>
        ///     Returns if this class is currently chaining Word UndoRecords.
        /// </summary>
        
        
        
        public bool IsRecording => _level > 0;


        /// <summary>
        ///     The name of the Word.UndoRecord entries currently being recorded.
        /// </summary>
        
        
        /// <remarks>
        ///     If the entries are Suspended then the entries that appear in the Word Undo list will be numbers as such:
        ///     UndoRecord, UndoRecord (1) , UndoRecord (2) , UndoRecord (3) etc.
        /// </remarks>
        public string CustomRecordName { get; private set; }


        private static UndoRecord UndoRecord
            => UndoRecordsSupported ? WordExtensions.WordApplication.Application.UndoRecordLJ() : null;


        private static bool UndoRecordsSupported => WordExtensions.WordVersion >= OfficeVersion.Office2010;


        /// <summary>
        ///     Inserts an UndoMarker, When EndUndoMarker is called the document is undone to the point at which UndoMarker was
        ///     called
        /// </summary>
        /// <param name="document">The document to place the undo marker in.</param>
        /// <remarks>
        ///     It is important that the calls to StartUndoMarker and EndUndoMarker are balanced.
        ///     If StartCustomRecord has been called then the Recording is automatically suspended and resumed when EndUndoMarker
        ///     is called.
        /// </remarks>
        public void StartUndoMarker(Document document)
        {
            Check.NotNull(document, "document");

            var markerState = _undoMarkerState;
            if (IsRecording && IsSuspended == false)
            {
                if (document == _document)
                {
                    SuspendRecording();
                    markerState = 0x2;
                }
            }

            //If _undoMarkerState = 0 Then
            //	If UndoRecordsSupported AndAlso document.ActiveWindow IsNot Nothing AndAlso document Is document.Application.ActiveDocumentLJ Then
            //		markerState = markerState Or &H4
            //		document.Application.UndoRecordLJ.StartCustomRecord(name)
            //	End If

            //End If
            _undoMarkerState = markerState;

            if (_undoMarkerState == 0)
            {
                _undoMarkerState |= 0x1;
            }

            UpdateUndoLevel(document, increment: true);
            _undoMarkerLevel.Lock();
        }


        /// <summary>
        ///     Called to end an UndoMarker operation started by calling StartUndoMarker.
        ///     This call calls Word.Document.Undo until it reaches the marker added by the StartUndoMarker call.
        /// </summary>
        /// <param name="document">The same document used in the call to StartUndoMarker</param>
        /// <param name="canRedo">If False will clear the redo stack so the operation cannot be re-applied.</param>
        
        public void EndUndoMarker(Document document, bool canRedo = true)
        {
            Check.NotNull(document, "document");

            if (_undoMarkerState == 0)
            {
                Debug.Assert(condition: false, message: "EndUndoMarker called without calling StartUndoMarker");
                // ReSharper disable once HeuristicUnreachableCode
                return;
            }




            Style style;
            var styles = document.Styles;

            if (styles.TryGetItemLJ(UndoStyleName, out style))
            {
                if (UndoRecordsSupported && (_undoMarkerState & 0x4) == 0x4)
                {
                    WordExtensions.WordApplication.UndoRecordLJ().EndCustomRecord();
                }

                var font = style.Font;
                var shading = font.Shading;
                var level = Convert.ToInt32(shading.BackgroundPatternColor);
                var undoSucceeded = true;
                while (!(undoSucceeded == false || (int) shading.BackgroundPatternColor < level))
                {
                    undoSucceeded = document.Undo();




                }

                if (undoSucceeded == false)
                {


                    Debug.Assert(condition: false, message: "undoSucceeded = False");
                }

                Marshal.ReleaseComObject(shading);
                Marshal.ReleaseComObject(font);
                Marshal.ReleaseComObject(style);
            }
            else
            {

                Debug.Assert(condition: false, message: "Cannot get style: " + UndoStyleName);
            }

            Marshal.ReleaseComObject(styles);

            if (canRedo == false)
            {
                ClearRedoStack(document, "Undo Marker");
            }

            if (_undoMarkerLevel.Unlock())
            {
                if ((_undoMarkerState & 0x2) == 0x2)
                {
                    _undoMarkerState = 0;
                    ResumeRecording();
                }

                _undoMarkerState = 0;
            }
        }


        /// <summary>
        ///     Creates an Word UndoRecord that can be suspended and resumed.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="endExistingUndoRecording"></param>
        /// <exception cref="InvalidOperationException">
        ///     InvalidOperationException is thrown if called during a StartUndoMarker operation.
        /// </exception>
        /// <remarks>
        ///     Word UndoRecord objects automatically end when the document changes and for some Object model calls.
        ///     Using StartCustomRecord/EndCustomRecord/SuspendCustomRecord/ResumeCustomRecord allows Word UndoRecord records to
        ///     continue to record.
        ///     This method cannot be called in-between StartUndoMarker/EndUndoMarker calls, doing
        ///     The number of calls to StartCustomRecord must be balanced by the same number of calls to EndCustomRecord
        ///     StartCustomRecords calls can be nested.
        ///     Word UndoRecords are only applicable to Word Versions 14 (2010) and greater, however the methods are safe to call
        ///     in all versions. The only
        ///     difference being that the Word Actions are not grouped into UndoRecords.
        /// </remarks>
        public int StartCustomRecord(string name = null, bool endExistingUndoRecording = false)
        {
            if (_undoMarkerState > 0)
            {
                throw new InvalidOperationException(
                                                    "Cannot call StartCustomRecord between StartUndoMarker EndUndoMarker calls");
            }

            if (string.IsNullOrEmpty(name))
            {
                name = "LJUndoRecord";
            }

            var ur = UndoRecord;

            if (ur != null)
            {
                if (endExistingUndoRecording && ur.IsRecordingCustomRecord)
                {
                    ur.EndCustomRecord();
                    Reset();
                }
            }

            if (string.IsNullOrEmpty(CustomRecordName))
            {
                CustomRecordName = name;
                Id = Environment.TickCount;
            }

            if (ur != null)
            {
                WordExtensions.WordApplication.UndoRecordLJ().StartCustomRecord(name);
            }

            _document = WordExtensions.WordApplication.ActiveDocument;

            _level++;

            return Id;
        }


        /// <summary>
        ///     Ends the recording of Word UndoRecords.
        /// </summary>
        /// <exception cref="InvalidOperationException">
        ///     InvalidOperationException is thrown if it's called during a SuspendRecording/ResumeRecording operation.
        /// </exception>
        
        public void EndCustomRecord()
        {
            if (IsRecording == false)
            {
                return;
            }

            if (IsSuspended)
            {
                throw new InvalidOperationException("EndCustomRecord not allowed when WordUndoRecord is Suspended");
            }

            if (UndoRecordsSupported)
            {
                WordExtensions.WordApplication.UndoRecordLJ().EndCustomRecord();
            }

            _level--;
            if (_level == 0)
            {
                Reset();
            }
        }


        /// <summary>
        ///     Called in-between StartCustomRecord/EndCustomRecord calls to temporarily suspend Word UndoRecord recording.
        ///     This is typically done when you need to temporarily switch to another document.
        /// </summary>
        /// <remarks>
        ///     ResumeRecording must be called before EndCustomRecord. SuspendRecording must be balanced with the same number
        ///     of calls to ResumeRecording
        /// </remarks>
        public void SuspendRecording()
        {
            if (IsRecording == false)
            {
                return;
            }

            if (_suspended.Lock() == false) //Return if it's not the first lock
            {
                return;
            }

            var ur = UndoRecord;
            if (ur != null)
            {
                while (ur.IsRecordingCustomRecord)
                {
                    ur.EndCustomRecord();
                }
            }
        }


        /// <summary>
        ///     Resumes Word UndoRecord recording after calling SuspendRecording.
        /// </summary>
        
        public void ResumeRecording()
        {
            if (IsRecording == false)
            {
                return;
            }

            if (_suspended.Unlock() == false)
            {
                return;
            }

            EnsureRecordingCustomRecord();
        }


        /// <summary>
        ///     Call to ensure that Word's CustomRecord is continuing to record Undo entries.
        /// </summary>
        
        /// <remarks>
        ///     Word's UndoRecord objects can end recording if the active document changes or certain object model calls are
        ///     made. This method is automatically called when the Word.DocumentChange event is raised.
        /// </remarks>
        public bool EnsureRecordingCustomRecord()
        {
            if (IsSuspended || IsRecording == false)
            {
                return false;
            }

            if (WordExtensions.WordApplication.IsObjectValid[_document] == false)
            {
                Reset();
                return false;
            }

            if (_undoMarkerState > 0)
            {
                //We are in an StartUndoMarker/EndUndoMarker operation so cannot re-start at this time.
                return false;
            }

            if (_document != WordExtensions.WordApplication.ActiveDocument)
            {
                return false;
            }

            var ur = UndoRecord;
            if (ur != null)
            {
                var name = CustomRecordName + " (" + _partNumber + ")";

                if (ur.IsRecordingCustomRecord == false || ur.CustomRecordName != name)
                {
                    ur.EndCustomRecord();
                    _partNumber++;


                    name = CustomRecordName + " (" + _partNumber + ")";

                    // Call StartCustomRecord equal to the number of levels
                    // so that any EndCustomRecord calls continue to be balanced.

                    var level = _level;
                    while (level != 0)
                    {
                        ur.StartCustomRecord(name);
                        level--;
                    }
                }
            }

            return true;
        }


        /// <summary>
        ///     Inserts an entry in the Undo stack which clears the ability to call Redo.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="name"></param>
        
        public static void ClearRedoStack(Document document, string name)
        {
            var ur = WordExtensions.WordApplication.Application.UndoRecordLJ();
            var callEndCustomRecord = false;

            var activeDocument = WordExtensions.WordApplication.ActiveDocument;
            var isActiveDocument = document != activeDocument;

            if (ur != null && (isActiveDocument == false || ur.IsRecordingCustomRecord == false))
            {
                if (isActiveDocument == false)
                {
                    document.Activate();
                }

                if (!string.IsNullOrEmpty(name))
                {
                    name = name.Trim();
                }

                ur.StartCustomRecord(name);
                callEndCustomRecord = true;
            }

            try
            {
                var rng = document.Range();
                // Don't use Collapse' we need a range of at least 1
                // Note even a blank document has a range of 1.
                rng.End = rng.Start + 1;
                rng.Bold = rng.Bold;
            }
            finally
            {
                if (callEndCustomRecord)
                {
                    ur.EndCustomRecord();
                }
            }

            document.Undo();

            if (isActiveDocument == false)
            {
                activeDocument.Activate();
                WordExtensions.UndoRecord.EnsureRecordingCustomRecord();
            }
        }


        private void Reset()
        {
            _undoMarkerState = 0;
            CustomRecordName = null;
            _partNumber = 0;
            _level = 0;
            _document = null;
            _suspended.Reset();
            _undoMarkerLevel.Reset();
            Id = 0;
        }


        /// <summary>
        ///     Used by StartUndoMarker/EndUndoMarker calls to add an undo-able marker to the document, which is invisible to the
        ///     user.
        /// </summary>
        private static void UpdateUndoLevel(Document document, bool increment)
        {
            int styleIndex;
            Shading shading;
            Style style;
            Font font;
            var styles = document.Styles;

            if (styles.TryGetItemLJ(UndoStyleName, out style) == false)
            {
                if (increment == false)
                    return;

                style = styles.Add(UndoStyleName);
                style.Hidden = true;
                style.Visibility = false;

                //Can only use style.Font.Shading because
                // style.Shading.BackgroundPatternColor & style.Borders(x), style.Frame.x etc. cause style
                // undo entries to be added even when simply reading the values!
                // style.Borders.InsideColor only works when the document is Active.

                font = style.Font;
                shading = font.Shading;

                styleIndex = 0;
            }
            else
            {
                font = style.Font;
                shading = font.Shading;
                styleIndex = (int) shading.BackgroundPatternColor;
                if (styleIndex < 0)
                    styleIndex = 0;
            }

            styleIndex = increment ? styleIndex + 1 : styleIndex - 1;
            shading.BackgroundPatternColor = (WdColor) styleIndex;

            Marshal.ReleaseComObject(shading);
            Marshal.ReleaseComObject(font);
            Marshal.ReleaseComObject(style);
            Marshal.ReleaseComObject(styles);
        }
    }
}