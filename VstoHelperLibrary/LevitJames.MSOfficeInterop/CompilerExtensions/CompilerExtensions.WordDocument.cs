// © Copyright 2018 Levit & James, Inc.

using JetBrains.Annotations;
using LevitJames.Core;
using LevitJames.MSOffice.Internal;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace LevitJames.MSOffice.MSWord
{
    public static partial class Extensions
    {


        /// <summary>
        /// Returns if a document has more than one co-author
        /// </summary>
        /// <param name="source">Word document to get the full name.</param>
        /// <returns>true if the document has more than one co-author;false otherwise</returns>
        public static bool HasCoAuthors([NotNull] this Document source)
        {
            if (source.Application.VersionLJ() < OfficeVersion.Office2010)
                return false;

            dynamic doc14 = source;
            dynamic authors = doc14.CoAuthoring.Authors;
            return authors.Count > 1;

        }

        /// <summary>Gets the full name of a Word document, including if the document is ODMA-compliant.</summary>
        /// <param name="source">Word document to get the full name.</param>
        /// <returns>The path to the document.</returns>
        /// <remarks>
        ///     If the Word document has never been saved, i.e., "DocumentN",
        ///     where "N" is an integer, then "DocumentN" will be returned.
        /// </remarks>
        // ReSharper disable once InconsistentNaming
        public static string FullNameLJ([NotNull] this Document source)
        {
            Check.NotNull(source, "source");

            var wordApp = source.Application;
            var fullName = source.FullName;

            //With open sharepoint Document:
            // fullName is a Url
            // wordApp.Documents.Exists(fullName) will return false.
            // Note this behaviour may have changed and the Exists call may now return true in newer versions of word.

            //With open OneDrive Document:
            // fullName is a Url
            // wordApp.Documents.Exists(fullName) will return true.
            //  in this case we need to use System.IO.File.Exists which will return false.

            if (!wordApp.Documents.Exists(fullName) || !System.IO.File.Exists(fullName))
            {
                // May be an ODMA document
                var activeDoc = wordApp.ActiveDocumentLJ();

                if (source != activeDoc)
                {
                    // WordBasic.FileName retrieves the actual path only for the active document
                    source.Activate();
                }

                fullName = wordApp.WordBasicLJ().FileName;

                activeDoc?.Activate();
            }

            return fullName;

        }


        /// <summary>
        ///     Returns the value of the Document.Final property. Returns Unknown if the call is not available at this time.
        /// </summary>
        /// <param name="source">A Document instance.</param>
        /// <returns>One of the WordBoolean values.</returns>
        // ReSharper disable once InconsistentNaming
        public static WordBoolean FinalLJ([NotNull] this Document source)
        {
            Check.NotNull(source, nameof(source));

            if (source.Application.VersionLJ() < OfficeVersion.Office2007)
                return WordBoolean.False;

            bool final;
            // ReSharper disable once SuspiciousTypeConversion.Global
            if (GetDocument12(source).Final_Get(out final) == PropertyNotAvailable)
            {
                return WordBoolean.Unknown;
            }

            return final ? WordBoolean.True : WordBoolean.False;
        }


        /// <summary>
        ///     Returns the CompatibilityMode introduced in Word 2012. If using Word 2010 it returns a value of 12.
        /// </summary>
        /// <param name="source">A Document instance.</param>
        /// <returns>12 for Word 2007; The actual value returned from CompatibilityMode for all versions of Word after 2007</returns>
        // ReSharper disable once InconsistentNaming
        public static int CompatibilityModeLJ14([NotNull] this Document source)
        {
            Check.NotNull(source, nameof(source));

            return source.Application.VersionLJ() >= OfficeVersion.Office2007 ? GetDocument14(source).CompatibilityMode : 12;
        }


        /// <summary>
        ///     Returns the value of the Document.TrackFormatting property. Returns Unknown if the call is not available at this
        ///     time.
        /// </summary>
        /// <param name="source">A Document instance.</param>

        // ReSharper disable once InconsistentNaming
        public static WordBoolean TrackFormattingLJ([NotNull] this Document source)
        {
            Check.NotNull(source, nameof(source));

            if (source.Application.VersionLJ() < OfficeVersion.Office2007) return WordBoolean.False;

            bool value;
            if (GetDocument12(source).TrackFormatting_Get(out value) == 0)
            {
                return value ? WordBoolean.True : WordBoolean.False;
            }

            return WordBoolean.Unknown;
        }

        /// <summary>
        ///     Sets the value of the Document.TrackFormatting property.
        /// </summary>
        /// <param name="source">A Document instance.</param>
        /// <param name="value">The new value to set the TrackFormatting too.</param>
        // ReSharper disable once InconsistentNaming
        public static void TrackFormattingLJ([NotNull] this Document source, bool value)
        {
            Check.NotNull(source, nameof(source));

            GetDocument12(source).TrackFormatting_Let(value);
        }

        /// <summary>
        ///     Sets the value of the Document.TrackFormatting property using the WordBoolean, typically returned from the
        ///     TrackFormattingLJ call.
        /// </summary>
        /// <param name="source">A Document instance.</param>
        /// <param name="value">
        ///     The new value to set the TrackFormatting too. If the value is WordBoolean.Unknown then
        ///     TrackFormatting is not called.
        /// </param>
        public static void TrackFormattingLJ([NotNull] this Document source, WordBoolean value)
        {
            if (value == WordBoolean.Unknown)
                return;
            TrackFormattingLJ(source, value == WordBoolean.True);
        }


        /// <summary>
        ///     Returns the value of the Document.TrackRevisions property. Returns Unknown if the call is not available at this
        ///     time.
        /// </summary>
        /// <param name="source">A Document instance.</param>

        // ReSharper disable once InconsistentNaming
        public static WordBoolean TrackRevisionsLJ([NotNull] this Document source)
        {
            Check.NotNull(source, nameof(source));

            if (GetDocument12(source).TrackRevisions_Get(out var value) == 0)
            {
                return value ? WordBoolean.True : WordBoolean.False;
            }

            return WordBoolean.Unknown;
        }

        /// <summary>
        ///     Sets the value of the Document.TrackRevisions property, returning the previous value as a WordBoolean enum value.
        /// </summary>
        /// <param name="source">A Document instance.</param>
        /// <param name="value">The new value for the TrackRevisions.</param>
        // ReSharper disable once InconsistentNaming
        public static WordBoolean TrackRevisionsLJ(this Document source, bool value)
        {
            if (source == null)
                return WordBoolean.Unknown;

            var tr = TrackRevisionsLJ(source);
            if (tr != WordBoolean.Unknown && value != (tr == WordBoolean.True))
                GetDocument12(source).TrackRevisions_Let(value);

            return tr;
        }

        /// <summary>
        ///     Sets the value of the Document.TrackFormatting property using the WordBoolean, typically returned from the
        ///     TrackFormattingLJ call.
        /// </summary>
        /// <param name="source">A Document instance.</param>
        /// <param name="value">
        ///     The new value for the TrackRevisions. If the value is WordBoolean.Unknown then TrackFormatting is
        ///     not called.
        /// </param>
        public static void TrackRevisionsLJ([NotNull] this Document source, WordBoolean value)
        {
            if (value == WordBoolean.Unknown)
                return;
            TrackRevisionsLJ(source, value == WordBoolean.True);
        }


        /// <summary>
        ///     Determines if a document is user editable
        /// </summary>
        /// <param name="source">A Document instance.</param>

        /// <remarks>
        ///     This method determines if a document is editable by checking the following states
        ///     <para>1. Document.Final = False (Word 12 or greater)</para>
        ///     <para>
        ///         2. Document.ProtectionType/document.Content.GotoEditableRange. If null is returned from these then the
        ///         document has been protected from edits and there are no editable ranges available in the document.
        ///     </para>
        ///     If the document is marked as ReadOnly then it is still deemed editable.
        /// </remarks>
        // ReSharper disable once InconsistentNaming
        public static WordBoolean LJIsEditable([NotNull] this Document source)
        {
            Check.NotNull(source, nameof(source));

            var editable = WordBoolean.True;

            if (source.ProtectionType != WdProtectionType.wdNoProtection)
            {
                editable = WordBoolean.False;
            }
            else
            {
                var doc12 = GetDocument12(source);
                if (doc12.Final_Get(out var final) == PropertyNotAvailable)
                    editable = WordBoolean.Unknown;
                else
                    editable = final ? WordBoolean.False : WordBoolean.True;

                //Marshal.ReleaseComObject(doc12); //Do not release document objects
            }

            return editable;
        }


        private static Document12 GetDocument12(Document doc)
        {
            // ReSharper disable once SuspiciousTypeConversion.Global
            return (Document12)doc;
        }


        private static Document14 GetDocument14(Document doc)
        {
            // ReSharper disable once SuspiciousTypeConversion.Global
            return (Document14)doc;
        }


        /// <summary>
        ///     Returns the first range that is not Protected, i.e. the user can edit.
        /// </summary>
        /// <param name="source">A Document instance.</param>
        /// <returns>null if the whole document is protected.</returns>

        // ReSharper disable once InconsistentNaming
        public static Range LJGetFirstEditableInsertionPoint([NotNull] this Document source)
        {
            Check.NotNull(source, nameof(source));
            return source.LJGetFirstEditableRange(collapseStart: true);
        }

        /// <summary>
        ///     Returns the first range that is not Protected, i.e. the user can edit.
        /// </summary>
        /// <param name="source">A Document instance.</param>
        /// <param name="collapseStart">Collapses the returned range to its start position.</param>
        /// <returns>null if the whole document is protected.</returns>

        // ReSharper disable once InconsistentNaming
        public static Range LJGetFirstEditableRange([NotNull] this Document source, bool collapseStart)
        {
            Check.NotNull(source, nameof(source));
            var rng = source.Range();

            if (source.ProtectionType != WdProtectionType.wdNoProtection)
            {
                rng.Collapse();
                rng = rng.GoToEditableRange();
            }

            if (rng == null) return null;

            if (collapseStart)
            {
                rng.Collapse();
            }

            return rng;
        }


        /// <summary>
        ///     Returns the last range that is not Protected, i.e. the user can edit.
        /// </summary>
        /// <param name="source">A Document instance.</param>
        /// <returns>null if the whole document is protected.</returns>

        // ReSharper disable once InconsistentNaming
        public static Range LJGetLastEditableRange([NotNull] this Document source)
        {
            Check.NotNull(source, nameof(source));
            return LJGetLastEditableRange(source, collapseEnd: false);
        }

        /// <summary>
        ///     Returns the last range that is not Protected, i.e. the user can edit.
        /// </summary>
        /// <param name="source">A Document instance.</param>
        /// <param name="collapseEnd">Collapses the returned range to its end position.</param>
        /// <returns>null if the whole document is protected.</returns>

        // ReSharper disable once InconsistentNaming
        public static Range LJGetLastEditableRange([NotNull] this Document source, bool collapseEnd)
        {
            Check.NotNull(source, nameof(source));

            var rng = source.Range();

            if (source.ProtectionType != WdProtectionType.wdNoProtection)
            {
                rng.Collapse();

                Range lastRange = null;
                do
                {
                    rng = rng.GoToEditableRange();
                    if (rng == null)
                    {
                        rng = lastRange;
                        break;
                    }

                    if (lastRange != null && rng.End <= lastRange.End)
                    {
                        rng = lastRange;
                        break;
                    }
                    if (lastRange != null)
                        Marshal.ReleaseComObject(lastRange);
                    lastRange = rng;
                } while (true);
            }

            if (rng == null || !collapseEnd) return rng;

            var rngStart = rng.Start; // failsafe marker

            rng.Collapse();
            //If Range ends in a paragraph mark then the end range will not get editable
            // so we must move back until it is editable
            bool editable;
            do
            {
                editable = LJIsEditable(source) == WordBoolean.True;
                if (editable || rng.Start <= rngStart)
                {
                    break;
                }

                rng.End = rng.End - 1;
            } while (true);

            if (editable == false)
            {
                Marshal.ReleaseComObject(rng);
                rng = null;
            }

            return rng;
        }


        /// <summary>
        ///     Determines if a document is editable, or is in one of a number of read-only states.
        ///     WordBoolean.Unknown is returned if the editable state cannot be determined at this time.
        /// </summary>
        /// <param name="source">A Document instance.</param>
        // ReSharper disable once InconsistentNaming
        public static WordBoolean LJIsProtectedView([NotNull] this Document source)
        {
            Check.NotNull(source, nameof(source));

            var isProtected = WordBoolean.False;

            if (source.Application.VersionLJ() < OfficeVersion.Office2007) return isProtected;

            // ReSharper disable once SuspiciousTypeConversion.Global
            // ReSharper disable once SuspiciousTypeConversion.Global
            var hr = ((Options12)source.Application.Options).AddBiDirectionalMarksWhenSavingTextFile_Get(out _);
            //****************************************************************************************
            // AddBiDirectionalMarksWhenSavingTextFile_Get will error on both the following cases.
            //1. If the document is "Protected" or Final
            //2. If a word floating dialog is active.
            // We need to know if the error is because of case 2.
            //   So we must further check the Word.Document.ActiveWindow.View.ShowCropMarks property
            //   ShowCropMarks only Errors for case 2.
            //****************************************************************************************

            if (Convert.ToBoolean(hr == PropertyNotAvailable)) // Only interested in hResult not retVal
            {
                var wordWindow = source.ActiveWindow;
                if (wordWindow != null)
                {
                    // ReSharper disable once SuspiciousTypeConversion.Global
                    isProtected = ((View12)wordWindow.View).ShowCropMarks_Get(out _) == 0 ? WordBoolean.True : WordBoolean.Unknown;
                }
            }
            else
            {
                isProtected = WordBoolean.False;
            }

            return isProtected;
        }


        /// <summary>
        ///     Cleans a document of Document Variables and/or  Bookmarks
        /// </summary>
        /// <returns>The total number of Word Variables and Bookmarks removed</returns>
        /// <param name="source">A Document instance.</param>
        /// <param name="bookmarkSearchPattern">A reg-ex search pattern to filter specific bookmark names. Can be a null string.</param>
        /// <param name="variableSearchPattern">A reg-ex search pattern to filter specific variable names. Can be a null string.</param>
        /// <param name="includeHiddenBookmarks">true to include hidden bookmarks;false otherwise.</param>
        // ReSharper disable once InconsistentNaming
        public static int LJCleanDocument([NotNull] this Document source, string bookmarkSearchPattern = null,
                                          string variableSearchPattern = null, bool includeHiddenBookmarks = false)
        {
            Check.NotNull(source, nameof(source));

            if (string.IsNullOrWhiteSpace(bookmarkSearchPattern) && string.IsNullOrWhiteSpace(bookmarkSearchPattern))
            {
                throw new ArgumentException("Invalid Search Patterns passed.");
            }

            var counter = 0;
            if (!string.IsNullOrWhiteSpace(bookmarkSearchPattern))
            {
                counter = source.LJCleanBookmarks(bookmarkSearchPattern, includeHiddenBookmarks);
            }

            if (!string.IsNullOrWhiteSpace(variableSearchPattern))
            {
                counter += source.LJCleanVariables(variableSearchPattern);
            }

            return counter;
        }


        /// <summary>
        ///     Cleans a Word.Document of Variables using the System.Text.RegularExpressions.Regex search pattern provided
        /// </summary>
        /// <param name="source">A Document instance.</param>
        /// <param name="variableSearchPattern">A reg-ex search pattern to filter specific variable names. Can be a null string.</param>
        /// <returns>The number of Word Variables removed</returns>
        // ReSharper disable once InconsistentNaming
        public static int LJCleanVariables([NotNull] this Document source, string variableSearchPattern)
        {
            Check.NotNull(source, nameof(source));
            Check.NotEmpty(variableSearchPattern, nameof(variableSearchPattern));

            var counter = 0;
            if (string.IsNullOrWhiteSpace(variableSearchPattern)) return counter;

            var regEx = new Regex(variableSearchPattern);

            var variables = source.Variables;
            for (var i = variables.Count; i >= 1; i--) // Word collections are 1-based
            {
                var variable = variables[i];
                if (regEx.IsMatch(variable.Name))
                {
                    variable.Delete();
                    counter += 1;
                }

                Marshal.ReleaseComObject(variable);
            }

            Marshal.ReleaseComObject(variables);
            return counter;
        }


        /// <summary>
        ///     Cleans a Word.Document of Bookmarks using the System.Text.RegularExpressions.Regex search pattern provided
        /// </summary>
        /// <param name="source">A Document instance.</param>
        /// <param name="bookmarkSearchPattern">A reg-ex search pattern to filter specific bookmark names.</param>
        /// <param name="includeHiddenBookmarks">true to include hidden bookmarks;false otherwise.</param>
        /// <returns>The number of Word Bookmarks removed</returns>
        // ReSharper disable once InconsistentNaming
        public static int LJCleanBookmarks([NotNull] this Document source, string bookmarkSearchPattern,
                                           bool includeHiddenBookmarks)
        {
            Check.NotNull(source, nameof(source));
            Check.NotEmpty(bookmarkSearchPattern, nameof(bookmarkSearchPattern));

            var bookmarks = source.Bookmarks;
            var showHidden = bookmarks.ShowHidden;
            try
            {
                bookmarks.ShowHidden = !includeHiddenBookmarks;
                bookmarks.ShowHidden = includeHiddenBookmarks;

                var counter = 0;
                if (string.IsNullOrWhiteSpace(bookmarkSearchPattern))
                    return counter;

                var regEx = new Regex(bookmarkSearchPattern);

                for (var i = bookmarks.Count; i >= 1; i--) // Word collections are 1-based
                {
                    var bookmark = bookmarks[i];
                    var bookmarkName = bookmark.Name;
                    if (regEx.IsMatch(bookmarkName))
                    {
                        bookmark.Delete();
                        counter += 1;
                    }

                    Marshal.ReleaseComObject(bookmark);
                }

                return counter;
            }
            finally
            {
                bookmarks.ShowHidden = !showHidden;
                bookmarks.ShowHidden = showHidden;
                Marshal.ReleaseComObject(bookmarks);
            }
        }


        /// <summary>
        ///     Cleans a Word.Document's custom properties using the System.Text.RegularExpressions.Regex search pattern provided.
        /// </summary>
        /// <param name="source">A Document instance.</param>
        /// <param name="propertiesSearchPattern">A reg-ex search pattern to filter specific document properties names.</param>
        /// <returns>The number of Word Bookmarks removed</returns>
        // ReSharper disable once InconsistentNaming
        public static int LJCleanProperties([NotNull] this Document source, string propertiesSearchPattern)
        {
            Check.NotNull(source, nameof(source));
            Check.NotEmpty(propertiesSearchPattern, nameof(propertiesSearchPattern));

            var counter = 0;
            var regEx = new Regex(propertiesSearchPattern);
            var docProperties = CustomDocumentPropertiesLJ(source);
            for (var i = docProperties.Count; i >= 1; i--) // Word collections are 1-based
            {
                var docProp = docProperties[i];
                if (regEx.IsMatch(docProp.Name))
                {
                    docProp.Delete();
                    counter += 1;
                }

                docProp.Dispose();
            }
            docProperties.Dispose();

            return counter;
        }


        /// <summary>
        ///     Suspends Tracking of Revisions and/or Formatting
        /// </summary>
        /// <param name="source">A Document instance.</param>
        /// <param name="suspendTrackRevisions">
        ///     On in; this should be set to true to suspend the tracking of revisions. On out;
        ///     this member is set to the previous TrackRevisions value. This value can then be passed to LJResumeTracking to
        ///     resume tracking of revisions.
        /// </param>
        /// <param name="suspendTrackFormatting">
        ///     On in; this should be set to true to suspend the tracking of formatting. On out;
        ///     this member is set to the previous TrackFormatting value. This value can then be passed to LJResumeTracking to
        ///     resume tracking of formatting.
        /// </param>

        // ReSharper disable once InconsistentNaming
        public static void LJSuspendTracking([NotNull] this Document source, ref bool suspendTrackRevisions,
                                             ref bool suspendTrackFormatting)
        {
            Check.NotNull(source, nameof(source));

            if (suspendTrackRevisions)
            {
                suspendTrackRevisions = source.TrackRevisionsLJ() == WordBoolean.True;
                if (suspendTrackRevisions)
                    source.TrackRevisions = false;
            }

            if (!suspendTrackFormatting) return;

            suspendTrackFormatting = source.TrackFormattingLJ() == WordBoolean.True;
            if (suspendTrackFormatting)
                source.TrackFormattingLJ(value: false);
        }


        /// <summary>Returns all the styles matching the styleType provided.</summary>
        /// <param name="source">Word document containing the styles.</param>
        /// <param name="styleType">One of the WdStyleType enums to filter the return collection.</param>
        /// <returns>Returns the name of the style or null if the style does not exist.</returns>
        public static IEnumerable<Style> GetStyles([NotNull] this Document source, WdStyleType styleType) => source.Styles.Cast<Style>().Where(s => s.Type == styleType);


        /// <summary>Returns all the style names matching the styleType provided.</summary>
        /// <param name="source">Word document containing the styles.</param>
        /// <param name="styleType">One of the WdStyleType enums to filter the return collection.</param>
        /// <returns>Returns a collection of style names.</returns>
        public static IEnumerable<string> GetStyleNames([NotNull] this Document source, WdStyleType styleType) => source.Styles.Cast<Style>().Where(s => s.Type == styleType).Select(s => s.NameLocal);


        /// <summary>Gets the font name and size from the provided style.</summary>
        /// <param name="source">Word document containing the style.</param>
        /// <param name="fontName">Name of the style.</param>
        /// <param name="fontSize">Name of the style.</param>

        public static bool TryGetStyleFontNameAndSize([NotNull] this Document source, string styleName, out string fontName, out float fontSize)
        {
            CheckNotNull(source);
            var fontStyle = source.GetStyleFont(styleName);
            if (fontStyle != null)
            {
                fontName = fontStyle.Name;
                fontSize = fontStyle.Size;
                Marshal.ReleaseComObject(fontStyle);
                return true;
            }

            fontName = null;
            fontSize = 0;
            return false;
        }


        /// <summary>
        ///     Returns a DocumentPropertiesLJ collection that represents all the custom document properties for the specified
        ///     document.
        ///     The Word version of this method must be called using reflection. The DocumentPropertiesLJ is a wrapper for
        ///     this collection and takes care of calling the methods.
        /// </summary>
        /// <param name="source">A Document instance.</param>
        /// <returns>A DocumentPropertiesLJ instance.</returns>


        // ReSharper disable once InconsistentNaming
        public static DocumentPropertyCollection CustomDocumentPropertiesLJ([NotNull] this Document source)
        {
            Check.NotNull(source, nameof(source));
            var props = source.CustomDocumentProperties;
            var result = new DocumentPropertyCollection(source.CustomDocumentProperties);
            Marshal.ReleaseComObject(props);
            return result;
        }


        /// <summary>
        ///     Returns a document custom property value, or the passed default value if the property does not exist.
        /// </summary>
        /// <param name="source">A Document instance.</param>
        /// <param name="name">The name of the property value to retrieve</param>
        /// <param name="defaultValue">A default value to return if the property does not exist.</param>
        public static object GetCustomDocumentProperty([NotNull] this Document source, string name, object defaultValue = null)
        {
            Check.NotNull(source, nameof(source));
            Check.NotEmpty(name, nameof(name));
            var props = source.CustomDocumentProperties;
            var result = DocumentPropertyCollection.GetValue(props, name, defaultValue);
            Marshal.ReleaseComObject(props);
            return result;
        }


        public static void SetCustomDocumentProperty([NotNull] this Document source, string name, object value)
        {
            Check.NotNull(source, nameof(source));
            Check.NotEmpty(name, nameof(name));
            var props = source.CustomDocumentProperties;
            DocumentPropertyCollection.AddOrUpdate(props, name, value);
            Marshal.ReleaseComObject(props);
        }


        public static object RemoveCustomDocumentProperty([NotNull] this Document source, string name)
        {
            Check.NotNull(source, nameof(source));
            Check.NotEmpty(name, nameof(name));
            var props = source.CustomDocumentProperties;
            var result = DocumentPropertyCollection.Remove(source.CustomDocumentProperties, name);
            Marshal.ReleaseComObject(props);
            return result;
        }


        /// <summary>
        ///     Returns a DocumentPropertiesLJ collection that represents all the BuiltIn document properties for the specified
        ///     document.
        ///     The Word version of this method must be called using reflection. The DocumentPropertiesLJ is a wrapper for
        ///     this collection and takes care of calling the methods.
        /// </summary>
        /// <param name="source">A Document instance.</param>
        /// <returns>A DocumentPropertiesLJ instance.</returns>


        // ReSharper disable once InconsistentNaming
        public static DocumentPropertyCollection BuiltInDocumentPropertiesLJ([NotNull] this Document source)
        {
            Check.NotNull(source, nameof(source));
            return new DocumentPropertyCollection(source.BuiltInDocumentProperties);
        }


        /// <summary>
        ///     Provides per Document customization of menu bars, toolbars, and key bindings, without altering the Saved or Active
        ///     state of the document.
        /// </summary>
        /// <typeparam name="T">The Generic Type to return from the CustomizationContextAction action.</typeparam>
        /// <param name="source">A Document instance.</param>
        /// <param name="contextCallback">A callback used to do the work during the context change,i.e. adding/removing CommandBars</param>

        // ReSharper disable once InconsistentNaming
        public static T LJCustomizationContext<T>([NotNull] this Document source, CustomizationContextAction<T> contextCallback)
        {
            Check.NotNull(source, nameof(source));
            Check.NotNull(contextCallback, nameof(contextCallback));

            var docSaved = source.Saved;
            object context = null;
            var activeDocument = source.Application.ActiveDocumentLJ();

            try
            {
                if (source != activeDocument)
                    source.Activate();
                else
                    activeDocument = null;

                try
                {
                    context = source.Application.CustomizationContext;
                }
                catch (Exception)
                {
                    //Debug.Assert(condition: false, message: "Context may have been set to an old document that is now closed and invalid.");
                    context = source.Application.NormalTemplate;
                }

                source.Application.CustomizationContext = source;

                return contextCallback.Invoke(source, context);
            }
            finally
            {
                source.Application.CustomizationContext = context;

                activeDocument?.Activate();
                source.Saved = docSaved;
            }
        }

        /// <summary>
        ///     Provides per Document customization of menu bars, toolbars, and key bindings, without altering the Saved or Active
        ///     state of the document.
        /// </summary>
        /// <typeparam name="T">The Generic Type to return from the CustomizationContextAction action.</typeparam>
        /// <param name="source">A Document instance.</param>
        /// <param name="contextCallback">A callback used to do the work during the context change,i.e. adding/removing CommandBars</param>

        // ReSharper disable once InconsistentNaming
        public static void LJCustomizationContext([NotNull] this Document source, CustomizationContextAction contextCallback)
        {
            source.LJCustomizationContext<object>((d, p) =>
            {
                contextCallback.Invoke(d, p);
                return null;
            });
        }


        /// <summary>Gets the font name of a style.</summary>
        /// <param name="source">Word document containing the style.</param>
        /// <param name="styleName">Name of the style.</param>
        /// <returns>Returns the name of the style or null if the style does not exist.</returns>
        public static string GetStyleFontName([NotNull] this Document source, string styleName)
        {
            CheckNotNull(source);
            return source.GetStyleFont(styleName)?.Name;
        }


        /// <summary>Gets the font size of a style.</summary>
        /// <param name="source">Word document containing the style.</param>
        /// <param name="styleName">Name of the style.</param>
        /// <returns>Returns the font size of the style or 0 if the style does not exist.</returns>
        public static float GetStyleFontSize([NotNull] this Document source, string styleName)
        {
            CheckNotNull(source);
            return source.GetStyleFont(styleName)?.Size ?? 0F;
        }


        /// <summary>Gets the font of a style.</summary>
        /// <param name="source">Word document containing the style.</param>
        /// <param name="styleName">The name of the style.</param>
        /// <returns>The style font.</returns>
        public static Font GetStyleFont([NotNull] this Document source, string styleName)
        {
            CheckNotNull(source);
            if (source.Styles.TryGetItemLJ(styleName, out var tempStyle))
            {
                var font = tempStyle.Font;
                Marshal.ReleaseComObject(tempStyle);
                return font;
            }

            return null;
        }


        /// <summary>
        ///     Renames an Word Style using the oldName and newName values
        /// </summary>
        /// <param name="source">The Word.Document containing the Style to rename</param>
        /// <param name="oldName">The existing name of the Style to rename.</param>
        /// <param name="newName">The new name for the Style.</param>
        public static void RenameStyle([NotNull] this Document source, string oldName, string newName)
        {
            source.Application.OrganizerRename(source.FullName, oldName, newName, WdOrganizerObject.wdOrganizerObjectStyles);
        }


        /// <summary>Generates a collapsed range in the main story of a document.</summary>

        // ReSharper disable once InconsistentNaming
        public static Range RangeLJ([NotNull] this _Document source, int insertionPoint)
        {
            CheckNotNull(source);
            return source.RangeLJ(WdStoryType.wdMainTextStory, insertionPoint);
        }

        /// <summary>Generates a collapsed range in a specified story of a document.</summary>

        // ReSharper disable once InconsistentNaming
        public static Range RangeLJ([NotNull] this _Document source, WdStoryType storyType, int insertionPoint)
        {
            CheckNotNull(source);
            return source.RangeLJ(storyType, insertionPoint, insertionPoint);
        }

        /// <summary>Generates a range in a specified story of a document.</summary>
        /// <returns>Returns the generated range.</returns>

        // ReSharper disable once InconsistentNaming
        public static Range RangeLJ([NotNull] this _Document source, WdStoryType storyType, int startPoint, int endPoint)
        {
            CheckNotNull(source);

            // Get the story range; there should always be a value
            var storyDup = source.StoryRanges.Cast<Range>().First(storyTemp => storyTemp.StoryType == storyType);

            // Make sure the endpoint doesn't go past the end of the story.
            // TODO: KDP - What if the endPoint value is much larger than the end of the story range... maybe the input story type is wrong?
            endPoint = Math.Min(endPoint, storyDup.End);
            if (startPoint > endPoint) startPoint = endPoint;

            try
            {
                storyDup.SetRange(startPoint, endPoint);
                return storyDup;
            }
            catch (COMException ex) when (ex.ErrorCode == -2146823680) //(Out of Range)
            {
                //Reduce and retry
                endPoint -= 1;
                if (startPoint > endPoint) startPoint = endPoint;
                storyDup.SetRange(startPoint, endPoint);
                return storyDup;
            }
        }


        /// <summary>Determines if Word document has deleted section breaks.</summary>
        /// <param name="doc">Word document to analyze.</param>
        /// <returns>Returns true if document contains deleted section breaks.</returns>
        public static bool HasDeletedSectionBreaks(this Document source)
        {
            return source.Sections.Cast<Section>().Any(sec => sec.SectionBreakHasBeenDeletedLJ());
        }


        public static Range LJPresetLocation(this Document source, PresetDocumentLocation location)
        {
            if (source == null || location == PresetDocumentLocation.None)
            {
                //Debug.Assert(false, "Invalid parameters: RangePickerData.GoToLocation");
                return null;
            }

            var wordApp = source.Application;

            Range presetLocation = null;
            switch (location)
            {
                case PresetDocumentLocation.BeginningOfDocument:
                    return source.Range(Start: 0, End: 0);
                case PresetDocumentLocation.EndOfDocument:
                    presetLocation = source.Range().Duplicate;
                    presetLocation.Collapse(WdCollapseDirection.wdCollapseEnd);
                    return presetLocation;
                case PresetDocumentLocation.EndOfCurrentSection:
                    presetLocation = wordApp.Selection.Sections[Index: 1].RangeLJ().Duplicate;
                    presetLocation.Collapse(WdCollapseDirection.wdCollapseEnd);
                    presetLocation.Move(WdUnits.wdCharacter, -1);
                    return presetLocation;
                case PresetDocumentLocation.EndOfPreviousSection:
                    if (PreviousSectionExists(source))
                    {
                        var currentSectionIndex = wordApp.Selection.Sections.First.Index;
                        var prevSectionIndex = source.Sections[currentSectionIndex].RangeLJ().Sections[Index: 1].Index - 1;
                        presetLocation = source.Sections[prevSectionIndex].RangeLJ().Duplicate;
                        presetLocation.Collapse(WdCollapseDirection.wdCollapseEnd);
                        presetLocation.Move(WdUnits.wdCharacter, -1);
                    }

                    return presetLocation;
                case PresetDocumentLocation.EndOfNextSection:
                    if (NextSectionExists(source))
                    {
                        var currentSection = wordApp.Selection.Sections.First;
                        presetLocation = source.Sections[source.Sections[currentSection.Index].RangeLJ().Sections.Last.Index + 1].RangeLJ()
                                               .Duplicate;
                        presetLocation.Collapse(WdCollapseDirection.wdCollapseEnd);
                        presetLocation.Move(WdUnits.wdCharacter, -1);
                    }

                    return presetLocation;
                case PresetDocumentLocation.EndOfToc:
                    var toc = GetDocumentToc(source);
                    if (toc != null)
                    {
                        presetLocation = toc.Range.Duplicate;
                        presetLocation.Collapse(WdCollapseDirection.wdCollapseEnd);
                    }

                    return presetLocation;

                case PresetDocumentLocation.BeginningOfFirstArabicSection:
                    return GetBeginningOfFirstArabicSection(source);

                case PresetDocumentLocation.None:
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(location), location, message: null);
            }

            return null;
        }

        private static TableOfContents GetDocumentToc(Document source)
        {
            return source.TablesOfContents.Count > 0 ? source.TablesOfContents[Index: 1] : null;
        }

        private static bool NextSectionExists(Document source)
        {
            return source.Application.Selection.Sections[Index: 1].RangeLJ().Sections.Last.Index < source.Sections.Count;
        }

        private static bool PreviousSectionExists(Document source)
        {
            return source.Application.Selection.Sections[Index: 1].RangeLJ().Sections[Index: 1].Index > 1;
        }


        public static bool LJPresetLocationExists(this Document source, PresetDocumentLocation location)
        {
            switch (location)
            {
                case PresetDocumentLocation.None:
                    return false;
                case PresetDocumentLocation.BeginningOfDocument:
                    return true;
                case PresetDocumentLocation.EndOfDocument:
                    return true;
                case PresetDocumentLocation.EndOfPreviousSection:
                    return PreviousSectionExists(source);
                case PresetDocumentLocation.EndOfCurrentSection:
                    return true;
                case PresetDocumentLocation.EndOfNextSection:
                    return NextSectionExists(source);
                case PresetDocumentLocation.EndOfToc:
                    return GetDocumentToc(source) != null;
                case PresetDocumentLocation.BeginningOfFirstArabicSection:
                    return GetBeginningOfFirstArabicSection(source) != null;

                default:
                    throw new ArgumentOutOfRangeException(nameof(location), location, null);
            }
        }


        private static Range GetBeginningOfFirstArabicSection(Document source)
        {
            var secNum = GetEndOfFrontMatterSectionIndex(source);
            var frontMatterAtBeginningOfDoc = secNum > 0;
            var section = secNum > 0 ? source.Sections[secNum] : null;
            var sectionRange = section?.RangeLJ();
            var firstSection = sectionRange?.Sections.First;
            var lastSection = sectionRange?.Sections.First;

            var firstSecNum = firstSection?.Index ?? 1;
            var lastSecNum = lastSection?.Index ?? 1;

            if (section != null)
                Marshal.ReleaseComObject(section);

            if (firstSection != null)
                Marshal.ReleaseComObject(firstSection);

            if (lastSection != null)
                Marshal.ReleaseComObject(lastSection);

            if (frontMatterAtBeginningOfDoc && firstSecNum > 0 && lastSecNum < source.Sections.Count)
            {
                var sect = source.Sections[lastSecNum + 1];
                var rng = sect.Range;

                var range = rng.Duplicate;
                range.Collapse(WdCollapseDirection.wdCollapseStart);

                Marshal.ReleaseComObject(sect);
                Marshal.ReleaseComObject(rng);
                return range;
            }

            return null;
        }


        public static int GetEndOfFrontMatterSectionIndex(this Document source)
        {
            var firstArabicIndex = GetFirstArabicPageSectionIndex(source); // Returns 9999 if none found
            var lastRomanIndex = 0;
            var idx = 0;
            foreach (var sct in source.LJSectionsExceptDeleted())
            {
                idx += 1;
                if (idx > firstArabicIndex)
                    break;
                var footer = sct.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
                var pageNumbers = footer.PageNumbers;

                var numStyle = pageNumbers.NumberStyle;
                if (numStyle == WdPageNumberStyle.wdPageNumberStyleLowercaseRoman)
                    lastRomanIndex = sct.Index;

                Marshal.ReleaseComObject(pageNumbers);
                Marshal.ReleaseComObject(footer);
                Marshal.ReleaseComObject(sct);
            }

            return lastRomanIndex;
        }


        private static int GetFirstArabicPageSectionIndex(Document wordDocument)
        {
            var docNonDeletedSections = wordDocument.LJSectionsExceptDeleted();
            var firstArabicIndex = 9999;
            var skipFirstSection = false;
            var firstSection = docNonDeletedSections.First();
            var footer = firstSection.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
            var footerPageNumbers = footer.PageNumbers;


            if (footerPageNumbers.NumberStyle == WdPageNumberStyle.wdPageNumberStyleArabic)
            {
                // Skip the first section if the first section is just the title page
                var rng = firstSection.Range;
                var workRange = rng.Duplicate;
                workRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                workRange.Move(WdUnits.wdCharacter, -1);
                // Collapse actually puts the range at the beginning of Section 2.
                var lastPage = (int)workRange.Information(WdInformation.wdActiveEndPageNumber);
                skipFirstSection = lastPage == 1;
                Marshal.ReleaseComObject(rng);
                Marshal.ReleaseComObject(workRange);
            }

            var firstSectionIndex = firstSection.Index;
            Marshal.ReleaseComObject(footer);
            Marshal.ReleaseComObject(footerPageNumbers);
            Marshal.ReleaseComObject(firstSection);


            foreach (var sct in docNonDeletedSections)
            {
                if (skipFirstSection && sct.Index <= firstSectionIndex)
                {

                    Marshal.ReleaseComObject(sct);
                    continue;
                }

                var footers = sct.Footers;
                footer = footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
                footerPageNumbers = footer.PageNumbers;

                var numStyle = footerPageNumbers.NumberStyle;
                var sectIndex = sct.Index;

                Marshal.ReleaseComObject(footers);
                Marshal.ReleaseComObject(footer);
                Marshal.ReleaseComObject(footerPageNumbers);
                Marshal.ReleaseComObject(sct);

                if (numStyle != WdPageNumberStyle.wdPageNumberStyleArabic)
                    continue;

                firstArabicIndex = sectIndex;
                break;
            }

            return firstArabicIndex;
        }


        [CLSCompliant(false)]
        public static void CloseLJ(this Document source, bool dontPromptToSave = true)
        {
            // Add handler to ensure that doc is Saved when closing.
            Application wordApp = null;
            if (dontPromptToSave)
            {
                wordApp = source.Application;
                wordApp.DocumentBeforeClose -= BeforeCloseHandler;
                wordApp.DocumentBeforeClose += BeforeCloseHandler;
            }

            try
            {
                source.Close();
            }
            finally
            {
                if (dontPromptToSave)
                {
                    wordApp.DocumentBeforeClose -= BeforeCloseHandler;
                }
            }
        }


        /// <summary>
        ///     Handler to ensure that document is marked as Saved after other Addins have made modifications without setting the
        ///     Saved flag.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="cancel"></param>

        private static void BeforeCloseHandler(Document doc, ref bool cancel) => doc.Saved = true;


        public static void AddXmlPart(this Document source, string xmlText, string xmlNameSpace)
        {
            RemoveCustomXmlParts(source, xmlNameSpace);
            source.CustomXMLParts.Add(xmlText);
        }


        public static bool ContainsXmlParts(this Document source, string xmlNameSpace)
        {
            var cxp = source.CustomXMLParts;
            var docXmlParts = cxp.SelectByNamespace(xmlNameSpace);
            return docXmlParts != null && docXmlParts.Count > 0;
        }


        public static void RemoveCustomXmlParts(this Document source, string xmlNameSpace)
        {
            if (string.IsNullOrEmpty(xmlNameSpace))
                return;

            var docXmlParts = source.CustomXMLParts.SelectByNamespace(xmlNameSpace);
            foreach (dynamic existingPart in docXmlParts)
            {
                existingPart.Delete();
                Marshal.ReleaseComObject(existingPart);
            }
        }


        public static void ClearTAFields(this Document document)
        {
            // Need to get the TA fields for each story range that supports it
            // Getting the Document.Fields only gets the fields in the main story range
            foreach (Range story in document.StoryRanges)
            {
                switch (story.StoryTypeLJ())
                {
                    case WdStoryType.wdMainTextStory:
                    case WdStoryType.wdFootnotesStory:
                    case WdStoryType.wdEndnotesStory:
                        //var taFields = story.Fields.Cast<Field>().Where(f => f.Type == WdFieldType.wdFieldTOAEntry).ToList();
                        foreach (var taField in story.Fields.Cast<Field>())
                        {
                            if (taField.Type == WdFieldType.wdFieldTOAEntry)
                                taField.Delete();

                            Marshal.ReleaseComObject(taField);
                        }

                        break;
                    default:
                        // Not supported
                        break;
                }

                Marshal.ReleaseComObject(story);
            }
        }


        public static void HideTrackFormattingBalloons(this Document document)
        {
            // ReSharper disable once UnusedVariable
            var numRevisions = document.Revisions.Count;
            // Querying number of revisions turns off TrackFormatting balloons in 2007!!! See TSWA #2408.
        }


        public static void ClearNullHeaderFooters(this Document document)
        {
            var activeWindow = document.ActiveWindow;
            var activeView = activeWindow.View;

            if (activeView.SplitSpecial != WdSpecialPane.wdPaneNone)
                activeWindow.Panes[2].Close();

            activeView = activeWindow.ActivePane.View;

            if (activeView.Type == WdViewType.wdNormalView || activeView.Type == WdViewType.wdOutlineView)
                activeView.Type = WdViewType.wdPrintView;

            activeView.SeekView = WdSeekView.wdSeekMainDocument;

            const int loopMax = 10;
            var loopCount = 0;
            while (loopCount < loopMax)
            {
                loopCount += 1;
                var view = activeWindow.ActivePane.View;
                view.SeekView = WdSeekView.wdSeekCurrentPageHeader;
                view.SeekView = WdSeekView.wdSeekMainDocument;
                document.Application.ScreenRefresh();
                System.Windows.Forms.Application.DoEvents();
                if (!document.HasNullHeaderFooter())
                    break;
            }
        }


        public static bool HasNullHeaderFooter(this Document document)
        {
            var pages = document.ActiveWindow.Panes[1].Pages;
            var pageCount = pages.Count;  // For troubleshooting
            var curPage = 0;              // For troubleshooting

            while (true)
            {
                curPage++;
                if (curPage > pages.Count)
                    break;
                var pageHasHdr = false;
                var pageHasFtr = false;
                Page page;
                try
                {
                    page = pages[curPage];
                }
                catch (COMException ex) when (ex.ErrorCode == -2146822347)
                {
                    // If the document is doing odd/even pages and there's a section break/next page,
                    // Word will insert a blank page. However, when accessing the Pages collection for that blank
                    // page, Word will throw an error that the member of the collection does not exist. In this
                    // case, just skip and go to the next page.
                    continue;
                }
                var rects = page.Rectangles;

                curPage++; // For troubleshooting
                for (var i = 1; i <= rects.Count; i++)
                {
                    var rct = rects[i];
                    var rng = rct.RangeLJ();
                    try
                    {
                        switch (rng?.StoryType)
                        {
                            case null:
                                break;

                            case WdStoryType.wdPrimaryHeaderStory:
                            case WdStoryType.wdFirstPageHeaderStory:
                            case WdStoryType.wdEvenPagesHeaderStory:
                                pageHasHdr = true;
                                break;
                            case WdStoryType.wdPrimaryFooterStory:
                            case WdStoryType.wdFirstPageFooterStory:
                            case WdStoryType.wdEvenPagesFooterStory:
                                pageHasFtr = true;
                                break;
                        }

                        if (pageHasHdr && pageHasFtr)
                            break;
                    }
                    finally
                    {
                        if (rng != null)
                            Marshal.ReleaseComObject(rng);

                        if (rct != null)
                            Marshal.ReleaseComObject(rct);
                    }
                }

                Marshal.ReleaseComObject(page);

                if (!pageHasHdr || !pageHasFtr)
                    return true;

            }

            return false;
        }
    }
}