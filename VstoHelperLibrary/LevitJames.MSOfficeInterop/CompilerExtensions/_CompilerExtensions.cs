// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using JetBrains.Annotations;
using LevitJames.Core;
using LevitJames.MSOffice.Internal;
using Microsoft.Office.Interop.Word;

namespace LevitJames.MSOffice.MSWord
{
    //
    // Summary:
    //     An Enum representing the major Microsoft Word versions.
    /// <summary>
    ///     Represents the of Compiler extensions added to the various Word and Office object models.
    /// </summary>
    /// <remarks>
    ///     The naming convention of Extension of methods is as follows:
    ///     <para>
    ///         If the Extension method is designed to replace an existing Word member then the Extension members
    ///         Name will be the same as the original Word member name.
    ///         However, the Extension member will end with the LJ suffix.
    ///     </para>
    ///     <para>
    ///         Common .Net methods have been added to some classes. These methods do not exist on the
    ///         Word Object model, but because of their usefulness, they were added.
    ///         Such methods include ExistsLJ and TryGetItemLJ.
    ///     </para>
    ///     <para>
    ///         If an extension method requires a minimum version of Word to work the
    ///         Extension method name is also suffixed with the version number of
    ///         Word required. In the figure above, there is a member called ExistsLJ14.
    ///         This member provides functionality introduced in version 14 of Word.
    ///     </para>
    ///     <para>
    ///         If you try to call an Extension method for a newer version of word than the version of
    ///         Word in use the extension method will simply return false or null. No Exception is ever thrown.
    ///     </para>
    ///     <para>
    ///         Any extension methods that add completely new functionality are always named with a LJ *Prefix*,
    ///         such as LJDialogs. Currently there are no methods provided by the OfficeInterop Assembly that add
    ///         new functionality to Word, but other Assemblies such as the WordExtensions Assembly do.
    ///         The base Word model for the Word Extensions is Word 2007 (version 12)
    ///     </para>
    /// </remarks>
    public static partial class Extensions
    {
        private const int PropertyNotAvailable = -2146823683; //&h800A11FD


        /// <summary>
        ///     Checks if the value passed is not null.
        /// </summary>
        /// <typeparam name="T">The type of value to check.</typeparam>
        /// <param name="value">The value to check.</param>
        /// <param name="parameterName">The name of the parameter to check.</param>
        /// <exception cref="ArgumentNullException">Throws an ArgumentNullException if the value is null.</exception>
        private static void CheckNotNull<T>(T source) where T : class
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }
        }


        /// <summary>
        ///     Tries to return a specific Word.Story instance from the Word.Stories collection.
        ///     Unlike the Word Stories.Item member this method will not throw an exception if the Story does not exist.
        /// </summary>
        /// <param name="source">A StoryRanges instance.</param>
        /// <param name="index">The WdStoryType index to retrieve.</param>
        /// <param name="storyRange">The Word.Story requested, or Null if the Word.Story does not exist</param>
        /// <returns>true on success; false on failure.</returns>
        
        // ReSharper disable once InconsistentNaming
        public static bool TryGetItemLJ([NotNull] this StoryRanges source, WdStoryType index, out Range storyRange)
        {
            Check.NotNull(source, nameof(source));
            // ReSharper disable once SuspiciousTypeConversion.Global
            var hr = ((StoryRanges11) source).Item(index, out storyRange);
            return hr == 0;
        }


        /// <summary>
        ///     Returns an IEnumerable of Range objects that match the story type passed through the storyType parameters.
        /// </summary>
        /// <param name="source">A StoryRanges instance.</param>
        /// <param name="storyType">A list containing the StoryTypes to return</param>
        
        public static IEnumerable<Range> OfType([NotNull] this StoryRanges source, params WdStoryType[] storyType)
        {
            Check.NotNull(source, nameof(source));
            foreach (Range sr in source)
            {
                if (storyType.Contains(sr.StoryType))
                    yield return sr;
            }
        }


        //View

        /// <summary>
        ///     Returns if the WordDocument.ActiveWindow.View.Type is in either wdNormalView or wdPrintView mode
        /// </summary>
        public static bool InNormalOrPrintView([NotNull] this View source)
        {
            switch (source.Type)
            {
            case WdViewType.wdNormalView:
            case WdViewType.wdPrintView:
            case WdViewType.wdPrintPreview:
                return true;
            default:
                return false;
            }
        }


        //Word.Dialog late bound calls


        /// <summary>
        ///     Calls Dialog.Display and returns the result the result as a WordDialogResult enum.
        /// </summary>
        /// <param name="source">A Dialog instance.</param>
        /// <returns>One of the WordDialogResult values.</returns>
        // ReSharper disable once InconsistentNaming
        public static WordDialogResult DisplayLJ([NotNull] this Dialog source)
        {
            Check.NotNull(source, nameof(source));

            return (WordDialogResult) source.Display();
        }


        //View

        /// <summary>
        ///     Changes the View.Type to wdPrintView if required, calls ShowRevisionsAndComments, and the switches the View.Type
        ///     back if required.
        /// </summary>
        /// <param name="source">A View instance.</param>
        /// <param name="value">true to show revisions and comments; false to hide them.</param>
        /// <returns>true if the ShowRevisionsAndComments call was successful; false otherwise.</returns>
        public static bool ShowRevisionsAndCommentsLJ([NotNull] this View source, bool value)
        {
            Check.NotNull(source, nameof(source));

            // Cannot set ShowRevisionsAndComments in ReadingView
            var viewType = source.Type;
            if (viewType == WdViewType.wdReadingView)
                source.Type = WdViewType.wdPrintView;

            var sucesss = false;
            try
            {
                source.ShowRevisionsAndComments = value;
                sucesss = true;
            }
            catch (COMException ex) when (ex.ErrorCode == -2146823683)
            {
                //This method or property is not available because the object refers to a footnote, endnote or comment.
                //Ignore
            }

            if (source.Type != viewType)
                source.Type = viewType;

            return sucesss;
        }
    }
}