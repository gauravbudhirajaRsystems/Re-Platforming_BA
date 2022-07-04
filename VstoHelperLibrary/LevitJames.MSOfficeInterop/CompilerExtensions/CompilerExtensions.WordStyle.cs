//*************************************************
//* © 2022 Litera Corp. All Rights Reserved.
//*************************************************

using System;
using System.Runtime.InteropServices;
using JetBrains.Annotations;
using LevitJames.Core;
using LevitJames.MSOffice.Internal;
using Microsoft.Office.Interop.Word;

namespace LevitJames.MSOffice.MSWord
{
    public static partial class Extensions
    {
        /// <summary>
        ///     Copies the style over to the provided target document using the OrgansizerCopy method. If the style is a built in
        ///     style and not in use then it is made an in use style before using.
        /// </summary>
        /// <param name="sourceStyle">The source style to copy</param>
        /// <param name="targetDocument">The document to copy the new style to.</param>
        public static bool OrganizerCopy([NotNull] this Style sourceStyle, [NotNull] Document targetDocument)
        {
            Check.NotNull(sourceStyle, nameof(sourceStyle));
            Check.NotNull(targetDocument, nameof(targetDocument));

            if (sourceStyle.InUse == false) //sourceStyle.BuiltIn && 
            {
                // Accessing a trivial property forces style to be InUse
                // If that's not done, then the organizer copy will fail
                var _ = sourceStyle.LanguageID;
            }

            // ReSharper disable once SuspiciousTypeConversion.Global
            var w11 = (WordApplication11) sourceStyle.Application;
            var sourceFileName = ((Document) sourceStyle.Parent).FullNameLJ();
            var targetFileName = targetDocument.FullNameLJ();
            var hr = w11.OrganizerCopy(sourceFileName, targetFileName, sourceStyle.NameLocal, WdOrganizerObject.wdOrganizerObjectStyles);
 
            switch (hr)
            {
            case 4605: // Ignore -- section protected for forms (don't every retry)
                return false;
            case 4198:
                // Sometimes, we get two different types of styles w/ same name, & OrganizerCopy fails. Try to delete from target, and then re-copy.
                DeleteStyleFromDoc(targetDocument, sourceStyle);
                hr = w11.OrganizerCopy(sourceFileName, targetFileName, sourceStyle.NameLocal, WdOrganizerObject.wdOrganizerObjectStyles);
                break;
            case 5587:
            case 5609:
                // Ignore -- Word occasionally creates 2 styles w/ same name & we get these duplication messages so roll w/ the punches (don't ever retry)
                return false;
            }

            if (hr != 0)
                Marshal.ThrowExceptionForHR(hr);
            return true;
        }

        private static void DeleteStyleFromDoc(Document document, Style style)
        {
            var styleName = style.NameLocal;

            style.Delete();

            // Check for alias

            foreach (Style docStyle in document.Styles)
            {
                if (docStyle.Type != WdStyleType.wdStyleTypeParagraph)
                    continue;

                foreach (var aliasName in docStyle.NameLocal.Split(new[] {","}, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (string.Compare(aliasName, styleName, StringComparison.CurrentCultureIgnoreCase) == 0)
                    {
                        docStyle.Delete();
                        return;
                    }
                }
            }
        }


        /// <summary>Copies this styles values over to the provided style..</summary>
        /// <param name="source">Source style.</param>
        /// <param name="targetDocument">The document to copy the new style to.</param>
        /// <param name="newStyleName">Target style.</param>
        public static void CopyTo([NotNull] this Style sourceStyle, [NotNull] Document targetDocument, [NotNull] string newStyleName)
        {
            var targetStyle = targetDocument.Styles.ExistsLJ(newStyleName)
                                  ? targetDocument.Styles[newStyleName]
                                  : targetDocument.Styles.Add(newStyleName, sourceStyle.Type);

            CopyTo(sourceStyle, targetStyle);
        }

        /// <summary>Copies this styles values over to the provided style..</summary>
        /// <param name="source">Source style.</param>
        /// <param name="targetStyle">Target style.</param>
        public static void CopyTo([NotNull] this Style sourceStyle, [NotNull] Style targetStyle)
        {
            targetStyle.AutomaticallyUpdate = sourceStyle.AutomaticallyUpdate;
            targetStyle.BaseStyle(sourceStyle.BaseStyle());
            targetStyle.Font = sourceStyle.Font.Duplicate;
            targetStyle.NextParagraphStyle(sourceStyle.NextParagraphStyle());
            targetStyle.NoProofing = sourceStyle.NoProofing;
            targetStyle.NoSpaceBetweenParagraphsOfSameStyle = sourceStyle.NoSpaceBetweenParagraphsOfSameStyle;
            targetStyle.ParagraphFormat = sourceStyle.ParagraphFormat.Duplicate;
            targetStyle.Borders = sourceStyle.Borders;
            targetStyle.LanguageID = sourceStyle.LanguageID;
            targetStyle.LanguageIDFarEast = sourceStyle.LanguageIDFarEast;
            targetStyle.AutomaticallyUpdate = sourceStyle.AutomaticallyUpdate;
        }


        /// <summary>
        ///     Tries to return a specific Word.Style instance from the Word.Styles collection.
        ///     Unlike the Word Stories.Item member this method will not throw an exception if the Story does not exist.
        /// </summary>
        /// <param name="source">A Styles instance.</param>
        /// <param name="index">The index to retrieve.</param>
        /// <param name="style">The Word.Style requested, or Null if the Word.Style does not exist</param>
        /// <returns>true on success; false on failure.</returns>
        
        // ReSharper disable once InconsistentNaming
        public static bool TryGetItemLJ([NotNull] this Styles source, int index, out Style style)
        {
            Check.NotNull(source, nameof(source));
            style = GetItemStyleCoreLJ(source, index);
            return style != null;
        }

        /// <summary>
        ///     Tries to return a specific Word.Style instance from the Word.Styles collection.
        ///     Unlike the Word Stories.Item member this method will not throw an exception if the Story does not exist.
        /// </summary>
        /// <param name="source">A Styles instance.</param>
        /// <param name="nameLocal">The name of the Word.Style to retrieve.</param>
        /// <param name="style">The Word.Style requested, or Null if the Word.Style does not exist</param>
        /// <returns>true on success; false on failure.</returns>
        
        // ReSharper disable once InconsistentNaming
        public static bool TryGetItemLJ([NotNull] this Styles source, [NotNull] string nameLocal, out Style style)
        {
            Check.NotNull(source, nameof(source));
            Check.NotEmpty(nameLocal, nameof(nameLocal));

            style = GetItemStyleCoreLJ(source, nameLocal);
            return style != null;
        }


        /// <summary>
        ///     Returns True if the Word.Style at the passed index exists.
        /// </summary>
        /// <param name="source">A Styles instance.</param>
        /// <param name="index">The integer index of the Word.Style to return</param>
        /// <returns>true if the Word.Style exists and is valid; false otherwise.</returns>
        
        // ReSharper disable once InconsistentNaming
        public static bool ExistsLJ([NotNull] this Styles source, int index)
        {
            Check.NotNull(source, nameof(source));

            var style = GetItemStyleCoreLJ(source, index);
            if (style == null) return false;
            Marshal.ReleaseComObject(style);
            return true;
        }

        /// <summary>
        ///     Returns True if the Word.Style exists.
        /// </summary>
        /// <param name="source">A Styles instance.</param>
        /// <param name="nameLocal">The name of the Word.Style to return</param>
        /// <returns>true if the Word.Style exists and is valid; false otherwise.</returns>
        
        // ReSharper disable once InconsistentNaming
        public static bool ExistsLJ([NotNull] this Styles source, [NotNull] string nameLocal)
        {
            Check.NotNull(source, nameof(source));
            if (string.IsNullOrEmpty(nameLocal))
                return false;

            var style = GetItemStyleCoreLJ(source, nameLocal);
            if (style == null) return false;
            Marshal.ReleaseComObject(style);
            return true;
        }


        // ReSharper disable once InconsistentNaming
        private static Style GetItemStyleCoreLJ([NotNull] Styles source, [NotNull] object index)
        {
            Style var = null;
            // ReSharper disable once SuspiciousTypeConversion.Global
            var safeStyles = (Styles11) source;
            var hr = safeStyles.Item(ref index, ref var);
            if (hr == 0 && var != null)
            {
                return var;
            }

            return null;
        }


        /// <summary>
        ///     Returns the BaseStyle of the supplied source Style.
        /// </summary>
        /// <param name="source">A Style instance.</param>
        /// <returns>A base Style instance.</returns>
        public static Style BaseStyle([NotNull] this Style source)
        {
            Check.NotNull(source, nameof(source));

            // ReSharper disable once UseIndexedProperty
            return (Style) source.get_BaseStyle();
        }

        /// <summary>
        ///     Sets the BaseStyle of the supplied source style name.
        /// </summary>
        /// <param name="source">A Style instance.</param>
        /// <param name="styleName">The name of the base style.</param>
        public static void BaseStyle([NotNull] this Style source, [NotNull] string styleName)
        {
            Check.NotNull(source, nameof(source));
            Check.NotNull(styleName, nameof(styleName));

            // ReSharper disable once UseIndexedProperty
            source.set_BaseStyle(styleName);
        }

        /// <summary>
        ///     Sets the BaseStyle of the supplied source Style.
        /// </summary>
        /// <param name="source">A Style instance.</param>
        /// <param name="style">The new base Style.</param>
        public static void BaseStyle([NotNull] this Style source, Style style)
        {
            Check.NotNull(source, nameof(source));

            // ReSharper disable once UseIndexedProperty
            source.set_BaseStyle(style);
        }

        /// <summary>
        ///     Sets the BaseStyle of the supplied source Style.
        /// </summary>
        /// <param name="source">A Style instance.</param>
        /// <param name="style">One of the WdBuiltinStyle values.</param>
        public static void BaseStyle([NotNull] this Style source, WdBuiltinStyle style)
        {
            Check.NotNull(source, nameof(source));
            Check.Enum(style, nameof(style));

            // ReSharper disable once UseIndexedProperty
            source.set_BaseStyle(style);
        }


        /// <summary>
        ///     Represents a link between a paragraph and a character style.
        /// </summary>
        /// <param name="source">A Style instance.</param>
        /// <returns>A link Style instance.</returns>
        public static Style LinkStyle([NotNull] this Style source)
        {
            Check.NotNull(source, nameof(source));
            // ReSharper disable once UseIndexedProperty
            return (Style) source.get_LinkStyle();
        }

        /// <summary>
        ///     Sets a link between a paragraph and a character style.
        /// </summary>
        /// <param name="source">A Style instance.</param>
        /// <param name="styleName">The name of the linked style.</param>
        public static void LinkStyle([NotNull] this Style source, string styleName)
        {
            Check.NotNull(source, nameof(source));
            Check.NotEmpty(styleName, nameof(styleName));

            // ReSharper disable once UseIndexedProperty
            source.set_LinkStyle(styleName);
        }

        /// <summary>
        ///     Sets the BaseStyle of the supplied source Style.
        /// </summary>
        /// <param name="source">A Style instance.</param>
        /// <param name="style">An instance of the style to link.</param>
        public static void LinkStyle([NotNull] this Style source, Style style)
        {
            Check.NotNull(source, nameof(source));
            Check.NotNull(style, nameof(style));

            // ReSharper disable once UseIndexedProperty
            source.set_LinkStyle(style);
        }

        /// <summary>
        ///     Sets the BaseStyle of the supplied source Style.
        /// </summary>
        /// <param name="source">A Style instance.</param>
        /// <param name="style">One of the WdBuiltinStyle values.</param>
        public static void LinkStyle([NotNull] this Style source, WdBuiltinStyle style)
        {
            Check.NotNull(source, nameof(source));
            Check.Enum(style, nameof(style));

            // ReSharper disable once UseIndexedProperty
            source.set_LinkStyle(style);
        }


        /// <summary>
        ///     Returns the style to be applied automatically to a new paragraph inserted after a paragraph formatted with the
        ///     specified style.
        /// </summary>
        /// <param name="source">A Style instance.</param>
        public static Style NextParagraphStyle([NotNull] this Style source)
        {
            Check.NotNull(source, nameof(source));
            // ReSharper disable once UseIndexedProperty
            return (Style) source.get_NextParagraphStyle();
        }

        /// <summary>
        ///     Sets the style to be applied automatically to a new paragraph inserted after a paragraph formatted with the
        ///     specified style.
        /// </summary>
        /// <param name="source">A Style instance.</param>
        /// <param name="styleName">The name of the style to apply to the NextParagraphStyle.</param>
        public static void NextParagraphStyle([NotNull] this Style source, string styleName)
        {
            Check.NotNull(source, nameof(source));
            Check.NotEmpty(styleName, nameof(styleName));
            // ReSharper disable once UseIndexedProperty
            source.set_NextParagraphStyle(styleName);
        }

        /// <summary>
        ///     Sets the style to be applied automatically to a new paragraph inserted after a paragraph formatted with the
        ///     specified style.
        /// </summary>
        /// <param name="source">A Style instance.</param>
        /// <param name="style">One of the WdBuiltinStyle values.</param>
        public static void NextParagraphStyle([NotNull] this Style source, WdBuiltinStyle style)
        {
            Check.NotNull(source, nameof(source));
            Check.Enum(style, nameof(style));
            // ReSharper disable once UseIndexedProperty
            source.set_NextParagraphStyle(style);
        }

        /// <summary>
        ///     Sets the style to be applied automatically to a new paragraph inserted after a paragraph formatted with the
        ///     specified style.
        /// </summary>
        /// <param name="source">A Style instance.</param>
        /// <param name="style">An instance of the next paragraph style to set.</param>
        public static void NextParagraphStyle([NotNull] this Style source, Style style)
        {
            Check.NotNull(source, nameof(source));
            Check.NotNull(style, nameof(style));
            // ReSharper disable once UseIndexedProperty
            source.set_NextParagraphStyle(style);
        }


        /// <summary>
        /// Determines whether the style can/will be seen in the Styles gallery.
        /// </summary>
        /// <param name="source">A Word.Style instance.</param>
        /// <returns>True, if the style will be seen in the Styles gallery.</returns>
        /// <remarks>Added this extension because the Word API call for this is both confusing and unintuitive.</remarks>
        public static bool GetVisibleInStylesGallery([NotNull] this Style source)
        {
            Check.NotNull(source, nameof(source));
            return source.Visibility == false;      // Yes, it's seemingly backwards in the Word API
        }

        /// <summary>
        /// Sets style visibility in the Styles gallery.
        /// </summary>
        /// <param name="source">A Word.Style instance.</param>
        /// <param name="visible">True if Style visible in Styles gallery.</param>
        /// <remarks>Added this extension because the Word API call for this is both confusing and unintuitive.</remarks>
        public static void SetVisibleInStylesGallery([NotNull] this Style source, bool visible)
        {
            Check.NotNull(source, nameof(source));
            source.Visibility = !visible;
        }
    }
}