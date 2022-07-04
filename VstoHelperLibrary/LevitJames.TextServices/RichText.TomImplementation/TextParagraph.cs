// © Copyright 2018 Levit & James, Inc.

// ReSharper disable once RedundantUsingDirective
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace LevitJames.TextServices
{
#pragma warning disable CS0659 // Type overrides Object.Equals(object o) but does not override Object.GetHashCode()

    /// <summary>
    /// Contains the paragraph properties of a TextRange instance.
    /// </summary>
    /// <remarks>Open a web browser browser and search for ITextParagraph.xxxx for detailed documentation.</remarks>
    [DebuggerStepThrough]
    public class TextParagraph : ITextPara, IEquatable<TextParagraph>, IDisposable
#pragma warning restore CS0659 // Type overrides Object.Equals(object o) but does not override Object.GetHashCode()
    {
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private ITextPara _para;

#if (TRACK_DISPOSED)
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private readonly string _disposedSource;
#endif

        internal TextParagraph(ITextPara para)
        {
            _para = para;
#if (TRACK_DISPOSED)
            _disposedSource = LevitJames.Core.Diagnostics.DisposeTracker.GetSource();
#endif
        }

        /// <summary>
        ///     Creates a duplicate of the specified paragraph format object.
        /// </summary>

        /// <returns>A Duplicate TextParagraph instance.</returns>

        public TextParagraph Duplicate
        {
            get => new TextParagraph(_para.Duplicate);
            set => _para.Duplicate = value._para;
        }

        [DebuggerHidden]
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        ITextPara ITextPara.Duplicate
        {
            get => _para.Duplicate;
            set => _para.Duplicate = value;
        }

        /// <summary>
        ///     Determines whether the paragraph formatting can be changed
        /// </summary>
        /// <returns>True if the paragraph formatting can be changed or False if it cannot be changed.</returns>
        /// <remarks>
        ///     CanChange returns True only if the paragraph formatting can be changed
        ///     (that is, if no part of an associated range is protected and an associated document is not read-only).
        ///     If this TextParagraph object is a duplicate, no protection rules apply.
        /// </remarks>
        public bool CanChange()
        {
            return _para.CanChange();
        }

        TomBoolean ITextPara.IsEqual(ITextPara para) => _para.IsEqual(para);

        /// <summary>
        ///     Determines paragraph formatting to a choice of default values.
        /// </summary>
        /// <param name="value">Type of reset.</param>

        public void Reset(ResetParagraphValue value)
        {
            _para.Reset(value);
        }

        /// <summary>
        ///     Determines the paragraph style for the paragraphs in a range.
        /// </summary>
        /// <value>A BuiltInStyles enum value that represents the style to apply.</value>
        /// <returns>A BuiltInStyles enum value that represents the style applied.</returns>
        /// <remarks>
        ///     The Text Object Model (TOM) version 1.0 has no way to
        ///     specify the meanings of user-defined style handles.
        ///     They depend on other facilities of the text system implementing TOM.
        ///     Negative style handles are reserved for built-in character and paragraph styles.
        ///     Currently defined values are listed in the following table.
        ///     For a description of the following styles, see the Microsoft Word documentation.
        /// </remarks>
        public BuiltInStyles Style
        {
            get => _para.Style;
            set => _para.Style = value;
        }

        /// <summary>
        ///     Gets/sets the paragraph alignment value.
        /// </summary>
        /// <value>A new paragraph alignment value to apply.</value>
        /// <returns>A current paragraph alignment applied.</returns>

        public ParagraphAlignment Alignment
        {
            get => _para.Alignment;
            set => _para.Alignment = value;
        }

        /// <summary>
        ///     Determines whether automatic hyphenation is enabled for the range.
        /// </summary>

        /// <returns>
        ///     One of the following values.
        ///     TomBoolean.True Automatic hyphenation is enabled.
        ///     TomBoolean.False Automatic hyphenation is disabled.
        ///     TomBoolean.Undefined The hyphenation property is undefined.
        /// </returns>

        public TomBoolean Hyphenation
        {
            get => _para.Hyphenation;
            set => _para.Hyphenation = value;
        }

        /// <summary>
        ///     Determines the amount used to indent the first line of a paragraph relative to the left indent.
        ///     The left indent is the indent for all lines of the paragraph except the first line.
        /// </summary>

        /// <returns>The first-line indentation amount in floating-point points.</returns>
        /// <remarks>
        ///     To set the first line indentation amount, call the TextParagraph.Indents method.
        ///     To get and set the indent for all other lines of the paragraph (that is, the left indent),
        ///     use TextParagraph.LeftIndent and TextParagraph.Indents.
        /// </remarks>
        public float FirstLineIndent => _para.FirstLineIndent;

        /// <summary>
        ///     Determines whether page breaks are allowed within paragraphs.
        /// </summary>

        /// <returns>
        ///     One of the following values.
        ///     TomBoolean.True. Page breaks are not allowed within a paragraph.
        ///     TomBoolean.False. Page breaks are allowed within a paragraph.
        ///     TomBoolean.Undefined. The property is undefined.
        /// </returns>

        public TomBoolean KeepTogether
        {
            get => _para.KeepTogether;
            set => _para.KeepTogether = value;
        }

        /// <summary>
        ///     Determines whether page breaks are allowed between paragraphs in the range.
        /// </summary>

        /// <returns>
        ///     One of the following values.
        ///     TomBoolean.True. Page breaks are not allowed within a paragraph.
        ///     TomBoolean.False. Page breaks are allowed within a paragraph.
        ///     TomBoolean.Undefined. The property is undefined.
        /// </returns>

        public TomBoolean KeepWithNext
        {
            get => _para.KeepWithNext;
            set => _para.KeepWithNext = value;
        }

        /// <summary>
        ///     Determines the distance used to indent all lines except the first line of a paragraph.
        ///     The distance is relative to the left margin.
        /// </summary>
        /// <value>The left indentation, in floating-point points to apply.</value>
        /// <returns>The left indentation, in floating-point points applied.</returns>
        /// <remarks>
        ///     To set the left indentation amount, call the TextParagraph.Indents method.
        ///     To get the first-line indent, call TextParagraph.FirstLineIndent.
        /// </remarks>
        public float LeftIndent => _para.LeftIndent;

        /// <summary>
        ///     Determines the line-spacing value for the text range,
        /// </summary>
        /// <value>
        ///     The line-spacing value to apply. The table below shows how this value is interpreted for the different
        ///     line-spacing rules.
        /// </value>
        /// <returns>
        ///     The line-spacing value applied. The table below shows how this value is interpreted for the different
        ///     line-spacing rules.
        /// </returns>
        /// <remarks>
        ///     The Paragraph.LineSpacingRule property determines how this value is interpreted.
        /// </remarks>
        public float LineSpacing => _para.LeftIndent;

        /// <summary>
        ///     Gets or sets the line-spacing rule for the text range
        /// </summary>

        /// <returns>A line ParagraphLineSpacingRule enum value.</returns>

        public ParagraphLineSpacingRule LineSpacingRule => _para.LineSpacingRule;

        /// <summary>
        ///     Gets or sets the kind of alignment to use for bulleted and numbered lists.
        /// </summary>
        /// <value>A ParagraphListAlignment enum value indicating the kind of bullet and numbering alignment to apply.</value>
        /// <returns>A ParagraphListAlignment enum value indicating the kind of bullet and numbering alignment applied.</returns>
        /// <remarks>For a description of the different types of lists, see the TextParagraph.ListType method</remarks>
        public ParagraphListAlignment ListAlignment
        {
            get => _para.ListAlignment;
            set => _para.ListAlignment = value;
        }

        /// <summary>
        ///     Gets or sets the list level index used with paragraphs.
        /// </summary>
        /// <value>The new value to set.</value>

        /// <remarks>
        ///     <list>
        ///         <listheader>Value</listheader>
        ///         <item>0, No list.</item>
        ///         <item>1, First-level (outermost) list.</item>
        ///         <item>2, Second-level (nested) list. This is nested under a level 1 list item.</item>
        ///         <item>3, Third-level (nested) list. This is nested under a level 2 list item.</item>
        ///         <item>and so forth, Nesting continues similarly. Up to three levels are common in HTML documents.</item>
        ///     </list>
        /// </remarks>
        public int ListLevelIndex
        {
            get => _para.ListLevelIndex;
            set => _para.ListLevelIndex = value;
        }

        /// <summary>
        ///     Gets or sets the starting value or code of a list numbering sequence.
        /// </summary>
        /// <value>
        ///     The starting value or code of a list numbering sequence.
        ///     For the possible values, see the TextParagraph.ListType method.
        /// </value>


        public int ListStart
        {
            get => _para.ListStart;
            set => _para.ListStart = value;
        }

        /// <summary>
        ///     Gets or sets the list tab setting, which is the distance between the first-line indent and the text on the first
        ///     line.
        ///     The numbered or bulleted text is left-justified, centered, or right-justified at the first-line indent value.
        /// </summary>
        /// <value>The list tab setting. The list tab value is in floating-point points.</value>
        /// <returns>The list tab setting in floating-point points.</returns>
        /// <remarks>
        ///     To determine whether the numbered or bulleted text is left-justified, centered, or right-justified, call
        ///     TextParagraph.ListAlignment.
        /// </remarks>
        public float ListTab
        {
            get => _para.ListTab;
            set => _para.ListTab = value;
        }

        /// <summary>
        ///     Gets or sets the kind of numbering to use with paragraphs.
        /// </summary>
        /// <value>A ParagraphListType enum value indicating the kind of list numbering.</value>
        /// <returns>The ParagraphListType applied.</returns>
        /// <remarks>
        ///     Values above 32 correspond to Unicode values for bullets.
        ///     The following example numbers the paragraphs in a range, r, starting with the number 2 and following the numbers
        ///     with a period.
        ///     <example>
        ///         <code lang="vb">
        /// r.Paragraph.ListStart = 2
        /// r.Paragraph.ListType = ParagraphListType.NumberAsArabic Or ParagraphListType.Period
        /// </code>
        ///     </example>
        ///     For an example of ParagraphListType.ListNumberAsSequence, set ListStart = 0x2780, which gives you circled numbers.
        ///     The Unicode Standard has examples of many more numbering sequences.
        /// </remarks>
        public ParagraphListType ListType
        {
            get => _para.ListType;
            set => _para.ListType = value;
        }

        /// <summary>
        ///     Determines whether paragraph numbering is enabled.
        /// </summary>

        /// <returns>
        ///     One of the following values.
        ///     TomBoolean.True. Line numbering is disabled.
        ///     TomBoolean.False. Line numbering is enabled.
        ///     TomBoolean.Undefined. The property is undefined.
        /// </returns>
        /// <remarks>
        ///     Paragraph numbering is when the paragraphs of a range are numbered.
        ///     The number appears on the first line of a paragraph.
        /// </remarks>
        public TomBoolean NoLineNumber
        {
            get => _para.NoLineNumber;
            set => _para.NoLineNumber = value;
        }

        /// <summary>
        ///     Determines whether each paragraph in the range must begin on a new page.
        /// </summary>

        /// <returns>
        ///     One of the following values.
        ///     TomBoolean.True. Each paragraph in this range must begin on a new page.
        ///     TomBoolean.False. The paragraphs in this range do not need to begin on a new page.
        ///     TomBoolean.Undefined. The property is undefined.
        /// </returns>

        public TomBoolean PageBreakBefore
        {
            get => _para.PageBreakBefore;
            set => _para.PageBreakBefore = value;
        }

        /// <summary>
        ///     Determines the size of the right margin indent of a paragraph.
        /// </summary>
        /// <value>The indentation, in floating-point points to apply.</value>
        /// <returns>The indentation, in floating-point points applied.</returns>

        public float RightIndent
        {
            get => _para.RightIndent;
            set => _para.RightIndent = value;
        }

        /// <summary>
        ///     Sets the first-line indent, the left indent, and the right indent for a paragraph.
        /// </summary>
        /// <param name="startIndent">
        ///     Indent of the first line in a paragraph, relative to the left indent. The value is in
        ///     floating-point points and can be positive or negative.
        /// </param>
        /// <param name="leftIndent">
        ///     Left indent of all lines except the first line in a paragraph, relative to left margin. The
        ///     value is in floating-point points and can be positive or negative.
        /// </param>
        /// <param name="rightIndent">
        ///     Right indent of all lines in paragraph, relative to the right margin. The value is in
        ///     floating-point points and can be positive or negative. This value is optional.
        /// </param>
        /// <remarks>
        ///     Line indents are not allowed to position text in the margins.
        ///     If the first-line indent is set to a negative value (for an out-dented paragraph) while the left indent is zero,
        ///     the first-line indent is reset to zero. To avoid this problem while retaining property sets,
        ///     set the first-line indent value equal to zero either explicitly or by calling the TextParagraph.Reset method.
        ///     Then, call TextParagraph.SetIndents to set a nonnegative, left-indent value and set the desired first-line indent.
        /// </remarks>
        public void SetIndents(float startIndent, float leftIndent, float rightIndent = 0)
        {
            _para.SetIndents(startIndent, leftIndent, rightIndent);
        }

        /// <summary>
        ///     Sets the paragraph line-spacing rule and the line spacing for a paragraph.
        /// </summary>
        /// <param name="lineSpacingRule">Value of new line-spacing rule.</param>
        /// <param name="lineSpacing">
        ///     Value of new line spacing. If the line-spacing rule treats the Spacing value as a linear
        ///     dimension, then Spacing is given in floating-point points.
        /// </param>
        /// <remarks>
        ///     The line-spacing rule and line spacing work together, and as a result, they must be set together, much as the
        ///     first and left indents need to be set together.
        /// </remarks>
        public void SetLineSpacing(ParagraphLineSpacingRule lineSpacingRule, float lineSpacing)
        {
            _para.SetLineSpacing(lineSpacingRule, lineSpacing);
        }

        /// <summary>
        ///     Determines the amount of vertical space below a paragraph.
        /// </summary>
        /// <value>The space-after value, in floating-point points to apply.</value>
        /// <returns>The space-after value, in floating-point points applied.</returns>

        public float SpaceAfter
        {
            get => _para.SpaceAfter;
            set => _para.SpaceAfter = value;
        }

        /// <summary>
        ///     Determines the amount of vertical space above a paragraph.
        /// </summary>
        /// <value>The space-before value, in floating-point points to apply.</value>
        /// <returns>The space-before value, in floating-point points applied.</returns>

        public float SpaceBefore
        {
            get => _para.SpaceBefore;
            set => _para.SpaceBefore = value;
        }

        /// <summary>
        ///     Determines whether to control widows and orphans for the paragraphs in a range.
        /// </summary>

        /// <returns>
        ///     One of the following values.
        ///     TomBoolean.True. Prevents the printing of a widow or orphan.
        ///     TomBoolean.False. Allows the printing of a widow or orphan.
        ///     TomBoolean.Undefined. The property is undefined.
        /// </returns>
        /// <remarks>
        ///     A widow is created when the last line of a paragraph is printed by itself at the top of a page.
        ///     An orphan is when the first line of a paragraph is printed by itself at the bottom of a page.
        /// </remarks>
        public TomBoolean WidowControl
        {
            get => _para.WidowControl;
            set => _para.WidowControl = value;
        }

        /// <summary>
        ///     Retrieves the tab count.
        /// </summary>


        /// <remarks>
        ///     The tab count of a new instance can be nonzero, depending on the underlying text engine.
        ///     For example, Microsoft Word stories begin with no explicit tabs defined, while rich edit instances start with a
        ///     single explicit tab.
        ///     To be sure there are no explicit tabs (that is, to set the tab count to zero), call TextParagraph.ClearAllTabs.
        /// </remarks>
        int ITextPara.TabCount => _para.TabCount;

        private TextTabsCollection _tabs;
        /// <summary>
        ///     Retrieves a list of the TextTab objects in the paragraph.
        /// </summary>
        /// <returns>A list of TextTab objects containing position, alignment, and leader values for each tab in the paragraph.</returns>
        public TextTabsCollection Tabs => _tabs ?? (_tabs = new TextTabsCollection(_para));

        /// <summary>
        ///     Adds a tab at the displacement tabPos, with type tabAlign, and leader style, tabLeader.
        /// </summary>
        /// <param name="tabPos">New tab displacement, in floating-point points.</param>
        /// <param name="tabAlign">A ParagraphTabAlign options for the tab position.</param>
        /// <param name="tabLeader">
        ///     A ParagraphAddTabLeader enum value. A leader character is the character that is used to fill
        ///     the space taken by a tab character.
        /// </param>

        void ITextPara.AddTab(float tabPos, ParagraphTabAlignment tabAlign, ParagraphAddTabLeader tabLeader)
        {
            _para.AddTab(tabPos, tabAlign, tabLeader);
        }

        /// <summary>
        ///     Clears all tabs, reverting to equally spaced tabs with the default tab spacing.
        /// </summary>

        void ITextPara.ClearAllTabs()
        {
            _para.ClearAllTabs();
        }

        /// <summary>
        ///     Deletes a tab at a specified displacement.
        /// </summary>
        /// <param name="tabPos">Displacement, in floating-point points, at which a tab should be deleted.</param>

        void ITextPara.DeleteTab(float tabPos)
        {
            _para.DeleteTab(tabPos);
        }

        ///// <summary>
        /////     Retrieves tab parameters (displacement, alignment, and leader style) for a specified tab.
        ///// </summary>
        ///// <param name="tabIndexOrConstant">
        /////     Index of tab for which to retrieve info.
        /////     It can be either a numerical index or a ParagraphTabIndex enum value.
        /////     Since tab indexes are zero-based, iTab = zero gets the first tab defined, iTab = 1 gets the second tab defined, and
        /////     so forth.
        ///// </param>
        ///// <param name="tabPos">
        /////     A variable that receives the tab displacement, in floating-point points.
        /////     The return value of tabPos is zero if the tab does not exist and the return value of tabPos is TomBoolean.Undefined
        /////     if there are multiple values in the associated range.
        ///// </param>
        ///// <param name="tabAlign">Receives the tab alignment. For more information, see TextParagraph.AddTab. </param>
        ///// <param name="tabLeader">Receives the tab leader-character style. For more information, see TextParagraph.AddTab</param>

        void ITextPara.GetTab(ParagraphTabIndex tabIndexOrConstant, out float tabPos, out ParagraphTabAlignment tabAlign, out ParagraphAddTabLeader tabLeader)
        {
            _para.GetTab(tabIndexOrConstant, out tabPos, out tabAlign, out tabLeader);
        }


#pragma warning disable 659
        /// <summary>
        /// Determines if one TextParagraph is equal to another.
        /// </summary>
        /// <param name="obj"></param>

        public override bool Equals(object obj) => Equals(obj as TextParagraph);
#pragma warning restore 659

        /// <summary>
        ///     Determines if the current range has the same properties as a specified range.
        /// </summary>
        /// <param name="other">The TextParagraph range that is compared to the current range.</param>
        public bool Equals(TextParagraph other)
        {
            if (other == null)
                return false;
            return _para.IsEqual(other._para) == TomBoolean.True;
        }


        // ReSharper disable once UnusedParameter.Local
        private void Dispose(bool disposeDotNetObjects)
        {
            if (_para != null)
            {
                Marshal.ReleaseComObject(_para);
                _para = null;
            }
        }


        /// <inheritdoc />
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <inheritdoc />
        ~TextParagraph()
        {
#if (TRACK_DISPOSED)
            LevitJames.Core.Diagnostics.DisposeTracker.Add(_disposedSource);
#endif
            Dispose(false);
        }




    }
}