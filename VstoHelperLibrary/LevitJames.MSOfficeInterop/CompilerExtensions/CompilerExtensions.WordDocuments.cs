// © Copyright 2018 Levit & James, Inc.

using System;
using JetBrains.Annotations;
using LevitJames.Core;
using Microsoft.Office.Interop.Word;

namespace LevitJames.MSOffice.MSWord
{
    public static partial class Extensions
    {
        //Documents


        /// <summary>
        ///     Returns a Word.Document from the passed document name.
        ///     If the item does not exist the method will return false. If the Word.Document exists the
        ///     method will return True and pass the Word.Document to the document parameter.
        /// </summary>
        /// <param name="source">A Documents instance.</param>
        /// <param name="name">The name or full name of the Word.Document to return</param>
        /// <param name="document">
        ///     The parameter that will receive the Word.Document upon success. On failure this value is set to
        ///     null.
        /// </param>
        /// <returns>true on success; false on failure.</returns>
        
        public static bool TryGetItem([NotNull] this Documents source, [NotNull] string name, out Document document)
        {
            Check.NotNull(source, nameof(source));
            Check.NotEmpty(name, nameof(name));

            document = null;
            // ReSharper disable once SuspiciousTypeConversion.Global
            var hr = ((Internal.Documents) source).Item(name, ref document);
            return hr == 0;
        }


        /// <summary>
        ///     Tries to return a Word.Document from the passed index without throwing an exception.
        ///     If the item does not exist the method will return false. If the Word.Document exists the
        ///     method will return True and pass the Word.Document to the document parameter.
        /// </summary>
        /// <param name="source">A Documents instance.</param>
        /// <param name="name">The name or full name of the Word.Document to return</param>
        /// <returns>true on success; false on failure.</returns>
        public static bool Exists([NotNull] this Documents source, [NotNull] string name)
        {
            Document document;
            return TryGetItem(source, name, out document);
        }


        /// <summary>
        ///     Opens a Word document with additional options
        /// </summary>
        /// <param name="source">A Word Application instance</param>
        /// <param name="fileName">The file name of the document to open</param>
        /// <param name="addToRecentFiles">true to add to the Word recent documents list; false otherwise</param>
        /// <param name="openReadOnly">true to open the document as read only; false otherwise</param>
        /// <param name="openVisible">true to open the document as visible as read only; false otherwise</param>
        /// <param name="openTempCopy">
        ///     true to open the document as a temporary copy, if the user already has the document open.
        ///     This also stops the document from appearing in Windows Jump Lists.
        /// </param>
        /// <param name="disableAutoMacros">
        ///     true to disable Word Auto Macros from running when the document is opened;false
        ///     otherwise;
        /// </param>
        /// <returns>A Document instance.</returns>
        public static Document OpenDocument([NotNull] this Documents source, [NotNull] string fileName, bool addToRecentFiles = false, bool openReadOnly = false, bool openVisible = false,
                                            bool openTempCopy = false, bool disableAutoMacros = true)
        {
            return OpenDocumentCore(source.Application, fileName, openTempCopy, disableAutoMacros, addToRecentFiles, openReadOnly, openVisible);
        }

        /// <summary>
        ///     Opens a Word document with additional options
        /// </summary>
        /// <param name="source">A Word Application instance</param>
        /// <param name="fileName">The file name of the document to open</param>
        /// <param name="fileName">The file name of the document to open</param>
        /// <param name="openDocumentCallback">
        ///     A callback used to open the Word Document with whatever parameters are required.
        ///     Typically this is the Word.Application.Documents.Open call.
        /// </param>
        /// <param name="disableAutoMacros">
        ///     true to disable Word Auto Macros from running when the document is opened;false
        ///     otherwise;
        /// </param>
        /// <returns>A Document instance.</returns>
        public static Document OpenDocument([NotNull] this Documents source, [NotNull] string fileName, [NotNull] Func<Application, string, Document> openDocumentCallback, bool openTempCopy = false,
                                            bool disableAutoMacros = true)
        {
            Check.NotNull(openDocumentCallback, nameof(openDocumentCallback));
            return OpenDocumentCore(source.Application, fileName, openTempCopy, disableAutoMacros, openDocumentCallback: openDocumentCallback);
        }

        //Note when openDocumentCallback is used the addToRecentFiles, openVisible, openReadOnly arguments are ignored
        //Since they will be set in the callback
        private static Document OpenDocumentCore(Application source, string fileName, bool openTempCopy = false, bool disableAutoMacros = false, bool addToRecentFiles = false, bool openReadOnly = false,
                                                 bool openVisible = false, Func<Application, string, Document> openDocumentCallback = null)
        {
            Document openedDocument;

            var docFileName = fileName;
            if (openTempCopy)
                docFileName = new TemporaryFile(fileName, keepExtension: true).FileName;

            // If you set ConfirmConversions to False in the Documents.Open call, the user's Word Option
            // setting will change! Need to restore it after opening the document.
            var confirmConversions = source.Options.ConfirmConversions;

            if (disableAutoMacros)
                source.WordBasicLJ().DisableAutoMacros();

            var conversionMessageModified = false;
            try
            {
                switch (source.VersionLJ())
                {
                case OfficeVersion.Office2003:

                    // 2007 Compatibility pack for Office 2003 displays a conversion message when opening a .docx file
                    // ConfirmConversion parameter in the Documents.Open call and in Tools>Options do not suppress the message
                    // Must set a reg key value (or roll out a Group Policy, details unknown)
                    // Add reg key value, open the document, then close the reg key value
                    // Should be no permissions problems as it is in the HKCU hive
                    conversionMessageModified = source.ShowDocXConversionMessage(false);

                    if (openDocumentCallback != null)
                    {
                        openedDocument = openDocumentCallback(source, docFileName);
                        break;
                    }

                    openedDocument = source.Documents.Open(docFileName,
                                                           ConfirmConversions: false,
                                                           ReadOnly: openReadOnly,
                                                           AddToRecentFiles: addToRecentFiles,
                                                           Format: WdOpenFormat.wdOpenFormatAuto,
                                                           Visible: openVisible);

                    break;
                case OfficeVersion.Office2007:
                case OfficeVersion.Office2010:
                case OfficeVersion.Office2013:
                case OfficeVersion.Office2016:

                    if (openDocumentCallback != null)
                    {
                        openedDocument = openDocumentCallback(source, docFileName);
                        break;
                    }

                    openedDocument = source.Documents.Open(docFileName,
                                                           ConfirmConversions: false,
                                                           ReadOnly: openReadOnly,
                                                           AddToRecentFiles:
                                                           addToRecentFiles,
                                                           Format: WdOpenFormat.wdOpenFormatAuto,
                                                           Visible: openVisible);

                    break;
                default:
                    openedDocument = null;
                    break;
                }
            }
            finally
            {
                if (conversionMessageModified)
                    source.ShowDocXConversionMessage(true);

                if (confirmConversions)
                    source.Options.ConfirmConversions = true;

                if (disableAutoMacros)
                    source.WordBasicLJ().EnableAutoMacros();
            }

            return openedDocument;
        }


        ///// <summary>
        ///// Opens a Word document with additional options
        ///// </summary>
        ///// <param name="source">A Word Application instance</param>
        ///// <param name="fileName">The file name of the document to open</param>
        ///// <param name="addToRecentFiles">true to add to the Word recent documents list; false otherwise</param>
        ///// <param name="openReadOnly">true to open the document as read only; false otherwise</param>
        ///// <param name="openVisible">true to open the document as visible as read only; false otherwise</param>
        ///// <param name="openTempCopy">true to open the document as a temporary copy, if the user already has the document open. This also stops the document from appearing in Windows Jump Lists.</param>
        ///// <param name="disableAutoMacros">true to disable Word Auto Macros from running when the document is opened;false otherwise;</param>
        ///// <returns>A Document instance.</returns>
        //public static Document EnsureOpenDocument([NotNull] this Documents source, [NotNull] string fileName, bool addToRecentFiles, bool openReadOnly, bool openVisible, bool openTempCopy, bool disableAutoMacros = true)
        //{
        //    Check.FileExists(fileName, "fileName");

        //    if (source.TryGetItem(fileName, out Document document, checkFullName: true))
        //        return document;

        //    return OpenDocumentCore(source.Application, fileName, openTempCopy, disableAutoMacros, addToRecentFiles, openReadOnly, openVisible);
        //}

        ///// <summary>Looks for open document whose path matches input filePath. If none are open, opens document.</summary>
        ///// <param name="source">Word.Application object.</param>
        ///// <param name="filePath">File path of document.</param>
        ///// <param name="openTempCopy"></param>
        ///// <param name="openedDoc"></param>
        ///// <returns>Returns true if document is opened during this call.</returns>
        ///// <remarks>If filePath is null</remarks>
        //public static Document EnsureOpenDocument([NotNull] this Documents source, [NotNull] string fileName,[NotNull] Func<Application, string, Document> openDocumentCallback, bool openTempCopy, bool disableAutoMacros = true)
        //{
        //    Check.FileExists(fileName, "fileName");
        //    Check.NotNull(openDocumentCallback, nameof(openDocumentCallback));

        //    if (source.TryGetItem(fileName, out Document document, checkFullName: true))
        //        return document;

        //    return OpenDocumentCore(source.Application, fileName, openTempCopy, disableAutoMacros, openDocumentCallback: openDocumentCallback);

        //}
    }
}