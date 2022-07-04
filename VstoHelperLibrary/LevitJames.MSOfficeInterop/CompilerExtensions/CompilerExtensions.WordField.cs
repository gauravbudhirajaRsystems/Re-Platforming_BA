//*************************************************
//* © 2021 Litera Corp. All Rights Reserved.
//*************************************************

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using JetBrains.Annotations;
using LevitJames.Core;
using LevitJames.MSOffice.Internal;
using LevitJames.TextServices;
using Microsoft.Office.Interop.Word;

namespace LevitJames.MSOffice.MSWord
{
 
    public static partial class Extensions
    {

        ///<summary>Unlinks the field only if it is not of type WdFieldType.wdFieldSequence or WdFieldType.wdFieldIndexEntry</summary>
        ///<remarks> 
        ///Word's Field.Unlink method can hang if the Fields is of Type WdFieldType.wdFieldSequence or WdFieldType.wdFieldIndexEntry. These field Types cannot be unlinked according to the Word documentation.
        ///</remarks>
        public static bool UnlinkSafe([NotNull] this Field source)
        {
            CheckNotNull(source);

            switch (source.Type)
            {
                case WdFieldType.wdFieldSequence:
                case WdFieldType.wdFieldIndexEntry:
                    return true;
            }
 
            source.Unlink();
            return true;
        }
 
        /// <summary>Determines if a field is contained within a hyperlink field result.</summary>
        /// <returns>Returns true if the field is contained within a hyperlink field.</returns>
        public static bool IsInAHyperlinkResult([NotNull] this Field source)
        {
            CheckNotNull(source);
            return source.Code.HyperlinkResultContainer() != null;
        }


        /// <summary>Creates a dictionary of key/sourceText pairs corresponding to switches and their values.</summary>
        /// <returns>Returns a string dictionary containing the switches and values</returns>
        public static Dictionary<string, string> Switches([NotNull] this Field source)
        {
            CheckNotNull(source);

            // Build switch dictionary on cleaned version of code text
            var cleanText = source.Code.Text.Clean();
            if (string.IsNullOrEmpty(cleanText))
            {
                return null;
            }

            var elements = cleanText.Split('\\');

            var switches = new Dictionary<string, string>();
            foreach (var element in elements)
            {
                // First word is the switch key; after that is the sourceText
                var switchElements = element.Split(" ".ToCharArray(), count: 2); // Splits only on the first space
                var key = $@"\{switchElements[0].Trim()}";
                var value =
                    switchElements.Length == 1
                        ? string.Empty
                        : switchElements[1].Clean().Replace($@"{Convert.ToChar(34)}", string.Empty).Trim();

                switches.Add(key.ToLower(), value);
            }

            return switches;
        }


        public static bool TypeHasResult([NotNull] this Field source)
        {
            switch (source.Type)
            {
                case WdFieldType.wdFieldAdvance:
                case WdFieldType.wdFieldAsk:
                case WdFieldType.wdFieldAddin:
                case WdFieldType.wdFieldAutoTextList:
                case WdFieldType.wdFieldComments:
                case WdFieldType.wdFieldData:
                case WdFieldType.wdFieldHTMLActiveX:
                case WdFieldType.wdFieldImport:
                case WdFieldType.wdFieldIndexEntry:
                case WdFieldType.wdFieldMergeField:
                case WdFieldType.wdFieldMergeRec:
                case WdFieldType.wdFieldNext:
                case WdFieldType.wdFieldNextIf:
                case WdFieldType.wdFieldNoteRef:
                case WdFieldType.wdFieldEmbed:
                case WdFieldType.wdFieldEmpty:
                case WdFieldType.wdFieldRefDoc:
                case WdFieldType.wdFieldSection:
                case WdFieldType.wdFieldSectionPages:
                case WdFieldType.wdFieldTOAEntry:
                case WdFieldType.wdFieldTOCEntry:
                case WdFieldType.wdFieldOCX:
                case WdFieldType.wdFieldPrint:
                case WdFieldType.wdFieldPrivate:
                case WdFieldType.wdFieldQuote:
                case WdFieldType.wdFieldSet:
                    return false;

                case WdFieldType.wdFieldAddressBlock:
                case WdFieldType.wdFieldAuthor:
                case WdFieldType.wdFieldAutoNum:
                case WdFieldType.wdFieldAutoNumLegal:
                case WdFieldType.wdFieldAutoNumOutline:
                case WdFieldType.wdFieldAutoText:
                case WdFieldType.wdFieldBarCode:
                case WdFieldType.wdFieldBidiOutline:
                case WdFieldType.wdFieldCompare:
                case WdFieldType.wdFieldCreateDate:
                case WdFieldType.wdFieldDatabase:
                case WdFieldType.wdFieldDate:
                case WdFieldType.wdFieldDDE:
                case WdFieldType.wdFieldDDEAuto:
                case WdFieldType.wdFieldDocProperty:
                case WdFieldType.wdFieldDocVariable:
                case WdFieldType.wdFieldEditTime:
                case WdFieldType.wdFieldExpression:
                case WdFieldType.wdFieldFileName:
                case WdFieldType.wdFieldFileSize:
                case WdFieldType.wdFieldFillIn:
                case WdFieldType.wdFieldFootnoteRef:
                case WdFieldType.wdFieldFormCheckBox:
                case WdFieldType.wdFieldFormDropDown:
                case WdFieldType.wdFieldFormTextInput:
                case WdFieldType.wdFieldFormula:
                case WdFieldType.wdFieldGlossary:
                case WdFieldType.wdFieldGoToButton:
                case WdFieldType.wdFieldGreetingLine:
                case WdFieldType.wdFieldHyperlink:
                case WdFieldType.wdFieldIf:
                case WdFieldType.wdFieldInclude:
                case WdFieldType.wdFieldIncludePicture:
                case WdFieldType.wdFieldIncludeText:
                case WdFieldType.wdFieldIndex:
                case WdFieldType.wdFieldInfo:
                case WdFieldType.wdFieldKeyWord:
                case WdFieldType.wdFieldLastSavedBy:
                case WdFieldType.wdFieldLink:
                case WdFieldType.wdFieldListNum:
                case WdFieldType.wdFieldMacroButton:
                case WdFieldType.wdFieldMergeSeq:
                case WdFieldType.wdFieldNumChars:
                case WdFieldType.wdFieldNumPages:
                case WdFieldType.wdFieldNumWords:
                case WdFieldType.wdFieldPage:
                case WdFieldType.wdFieldPageRef:
                case WdFieldType.wdFieldPrintDate:
                case WdFieldType.wdFieldRef:
                case WdFieldType.wdFieldRevisionNum:
                case WdFieldType.wdFieldSaveDate:
                case WdFieldType.wdFieldSequence:
                case WdFieldType.wdFieldShape:
                case WdFieldType.wdFieldSkipIf:
                case WdFieldType.wdFieldStyleRef:
                case WdFieldType.wdFieldSubject:
                case WdFieldType.wdFieldSubscriber:
                case WdFieldType.wdFieldSymbol:
                case WdFieldType.wdFieldTemplate:
                case WdFieldType.wdFieldTime:
                case WdFieldType.wdFieldTitle:
                case WdFieldType.wdFieldTOA:
                case WdFieldType.wdFieldTOC:
                case WdFieldType.wdFieldUserAddress:
                case WdFieldType.wdFieldUserInitials:
                case WdFieldType.wdFieldUserName:
                case WdFieldType.wdFieldCitation:
                case WdFieldType.wdFieldBibliography:
                    return true;
                default:
                    return false;
            }
        }


        /// <summary>
        ///     Returns the Word.Range instance that belongs to a Word.Field instance.
        ///     Unlike the Word supplied Field.Range member this method will not throw an exception if the
        ///     Word.Range is not valid for the Field.
        /// </summary>
        /// <param name="source">A Field instance.</param>
        /// <returns>A Word.Range object if one exists for the field, Null otherwise.</returns>

        // ReSharper disable once InconsistentNaming
        public static Range ResultLJ([NotNull] this Field source)
        {
            Check.NotNull(source, nameof(source));

            Range range = null;
            // ReSharper disable once SuspiciousTypeConversion.Global
            var hr = ((Field11)source).Result(ref range);
            if (hr == 0)
            {
                return range;
            }

            return null;
        }
    }
}