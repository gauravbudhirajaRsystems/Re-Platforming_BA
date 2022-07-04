// © Copyright 2018 Levit & James, Inc.

using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace LevitJames.TextServices
{
    /// <summary>
    ///     Represents a Text Object Model interface ITextPara.
    /// </summary>
    /// <remarks>
    ///     The TextFont and TextParagraph interfaces encapsulate the functionality of the Microsoft Word Format Font and
    ///     Paragraph dialog boxes, respectively.
    ///     Both interfaces include a duplicate (Value) property that can return a duplicate of the attributes in a range
    ///     object or transfer a set of attributes to a range.
    ///     As such, they act like programmable format painters. For example, you could transfer all attributes from range r1
    ///     to range r2 except
    ///     for making r2 bold and the font size 12 points by using the following subroutine.
    ///     <example>
    ///         <code lang="vb">
    /// Sub AttributeCopy(ByVal r1 As TextRange, ByVal r2 As TextRange)
    ///  Dim tf As TextFont
    ///  tf = r1.Font                ' Value is the default property    
    ///  tf.Bold = BomBoolean.True   ' You can make some modifications
    ///  tf.Size = 12
    ///  tf.Animation = FontAnimation.SparkleText
    ///  r2.Font = tf                ' Apply font attributes all at once
    /// End Sub
    /// </code>
    ///     </example>
    ///     <para>
    ///         The TextParagraph interface encapsulates the Word Paragraph dialog box.
    ///         All measurements are given in floating-point points. The rich edit control is able to accept and return all
    ///         TextParagraph properties
    ///         intact (that is, without modification), both through TOM and through its Rich Text Format (RTF) converters.
    ///         However, the following properties have no effect on what the control displays:
    ///     </para>
    ///     <para>
    ///         <list>
    ///             <item>DoNotHyphen</item>
    ///             <item>KeepTogether</item>
    ///             <item>KeepWithNext</item>
    ///             <item>LineSpacing</item>
    ///             <item>LineSpacingRule</item>
    ///             <item>NoLineNumber</item>
    ///             <item>PageBreakBefore</item>
    ///             <item>Tab alignments</item>
    ///             <item>Tab styles (other than AlignLeft and Spaces)</item>
    ///             <item>Style WidowControl</item>
    ///         </list>
    ///     </para>
    ///     <para>
    ///         This is a Com interface so Marshal.ReleaseComObject should be used to free the object when you have finished
    ///         with it.
    ///     </para>
    /// </remarks>
    [CompilerGenerated]
    [SuppressMessage("Microsoft.Naming", "CA1715:IdentifiersShouldHaveCorrectPrefix", MessageId = "I")]
    [ComImport]
    [Guid("8CC497C4-A1DF-11CE-8098-00AA0047BE5D")]
    [TypeLibType(flags: 0x10C0)]
    [DefaultMember("Duplicate")]
    internal interface ITextPara
    {
        /// <summary>
        ///     Creates a duplicate of the specified paragraph format object.
        /// </summary>

        /// <returns>A Duplicate TextParagraph instance.</returns>

        [DispId(dispId: 0)]
        ITextPara Duplicate { get; set; }

        /// <summary>
        ///     Determines whether the paragraph formatting can be changed
        /// </summary>
        /// <returns>True if the paragraph formatting can be changed or False if it cannot be changed.</returns>
        /// <remarks>
        ///     CanChange returns True only if the paragraph formatting can be changed
        ///     (that is, if no part of an associated range is protected and an associated document is not read-only).
        ///     If this TextParagraph object is a duplicate, no protection rules apply.
        /// </remarks>
        [return: MarshalAs(UnmanagedType.I4)]
        [DispId(dispId: 0x401)]
        bool CanChange();

        /// <summary>
        ///     Determines if the current range has the same properties as a specified range.
        /// </summary>
        /// <param name="para">The TextParagraph range that is compared to the current range.</param>


        [return: MarshalAs(UnmanagedType.I4)]
        [DispId(dispId: 0x402)]
        TomBoolean IsEqual([In] [MarshalAs(UnmanagedType.Interface)]
                           ITextPara para);

        /// <summary>
        ///     Determines paragraph formatting to a choice of default values.
        /// </summary>
        /// <param name="value">Type of reset.</param>

        [DispId(dispId: 0x403)]
        void Reset([In] ResetParagraphValue value);

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
        [DispId(dispId: 0x404)]
        BuiltInStyles Style { get; set; }

        /// <summary>
        ///     Gets/sets the paragraph alignment value.
        /// </summary>
        /// <value>A new paragraph alignment value to apply.</value>
        /// <returns>A current paragraph alignment applied.</returns>

        [DispId(dispId: 0x405)]
        ParagraphAlignment Alignment { get; set; }

        /// <summary>
        ///     Determines whether automatic hyphenation is enabled for the range.
        /// </summary>

        /// <returns>
        ///     One of the following values.
        ///     TomBoolean.True Automatic hyphenation is enabled.
        ///     TomBoolean.False Automatic hyphenation is disabled.
        ///     TomBoolean.Undefined The hyphenation property is undefined.
        /// </returns>

        [DispId(dispId: 0x406)]
        TomBoolean Hyphenation { get; set; }

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
        [DispId(dispId: 0x407)]
        float FirstLineIndent { get; }

        /// <summary>
        ///     Determines whether page breaks are allowed within paragraphs.
        /// </summary>

        /// <returns>
        ///     One of the following values.
        ///     TomBoolean.True. Page breaks are not allowed within a paragraph.
        ///     TomBoolean.False. Page breaks are allowed within a paragraph.
        ///     TomBoolean.Undefined. The property is undefined.
        /// </returns>

        [DispId(dispId: 0x408)]
        TomBoolean KeepTogether { get; set; }

        /// <summary>
        ///     Determines whether page breaks are allowed between paragraphs in the range.
        /// </summary>

        /// <returns>
        ///     One of the following values.
        ///     TomBoolean.True. Page breaks are not allowed within a paragraph.
        ///     TomBoolean.False. Page breaks are allowed within a paragraph.
        ///     TomBoolean.Undefined. The property is undefined.
        /// </returns>

        [DispId(dispId: 0x409)]
        TomBoolean KeepWithNext { get; set; }

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
        [DispId(dispId: 0x410)]
        float LeftIndent { get; }


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
        [DispId(dispId: 0x411)]
        float LineSpacing { get; }

        /// <summary>
        ///     Gets or sets the line-spacing rule for the text range
        /// </summary>

        /// <returns>A line ParagraphLineSpacingRule enum value.</returns>

        [DispId(dispId: 0x412)]
        ParagraphLineSpacingRule LineSpacingRule { get; }


        /// <summary>
        ///     Gets or sets the kind of alignment to use for bullet-ed and numbered lists.
        /// </summary>
        /// <value>A ParagraphListAlignment enum value indicating the kind of bullet and numbering alignment to apply.</value>
        /// <returns>A ParagraphListAlignment enum value indicating the kind of bullet and numbering alignment applied.</returns>
        /// <remarks>For a description of the different types of lists, see the TextParagraph.ListType method</remarks>
        [DispId(dispId: 0x413)]
        ParagraphListAlignment ListAlignment { get; set; }

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
        [DispId(dispId: 0x414)]
        int ListLevelIndex { get; set; }

        /// <summary>
        ///     Gets or sets the starting value or code of a list numbering sequence.
        /// </summary>
        /// <value>
        ///     The starting value or code of a list numbering sequence.
        ///     For the possible values, see the TextParagraph.ListType method.
        /// </value>


        [DispId(dispId: 0x415)]
        int ListStart { get; set; }

        /// <summary>
        ///     Gets or sets the list tab setting, which is the distance between the first-line indent and the text on the first
        ///     line.
        ///     The numbered or bullet-ed text is left-justified, centered, or right-justified at the first-line indent value.
        /// </summary>
        /// <value>The list tab setting. The list tab value is in floating-point points.</value>
        /// <returns>The list tab setting in floating-point points.</returns>
        /// <remarks>
        ///     To determine whether the numbered or bullet-ed text is left-justified, centered, or right-justified, call
        ///     TextParagraph.ListAlignment.
        /// </remarks>
        [DispId(dispId: 0x416)]
        float ListTab { get; set; }

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
        [DispId(dispId: 0x417)]
        ParagraphListType ListType { get; set; }

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
        [DispId(dispId: 0x418)]
        TomBoolean NoLineNumber { get; set; }

        /// <summary>
        ///     Determines whether each paragraph in the range must begin on a new page.
        /// </summary>

        /// <returns>
        ///     One of the following values.
        ///     TomBoolean.True. Each paragraph in this range must begin on a new page.
        ///     TomBoolean.False. The paragraphs in this range do not need to begin on a new page.
        ///     TomBoolean.Undefined. The property is undefined.
        /// </returns>

        [DispId(dispId: 0x419)]
        TomBoolean PageBreakBefore { get; set; }

        /// <summary>
        ///     Determines the size of the right margin indent of a paragraph.
        /// </summary>
        /// <value>The indentation, in floating-point points to apply.</value>
        /// <returns>The indentation, in floating-point points applied.</returns>

        [DispId(dispId: 0x420)]
        float RightIndent { get; set; }

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
        [DispId(dispId: 0x421)]
        void SetIndents([In] float startIndent, [In] float leftIndent, [In] float rightIndent = 0F);

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
        [DispId(dispId: 0x422)]
        void SetLineSpacing([In] ParagraphLineSpacingRule lineSpacingRule, [In] float lineSpacing);

        /// <summary>
        ///     Determines the amount of vertical space below a paragraph.
        /// </summary>
        /// <value>The space-after value, in floating-point points to apply.</value>
        /// <returns>The space-after value, in floating-point points applied.</returns>

        [DispId(dispId: 0x423)]
        float SpaceAfter { get; set; }


        /// <summary>
        ///     Determines the amount of vertical space above a paragraph.
        /// </summary>
        /// <value>The space-before value, in floating-point points to apply.</value>
        /// <returns>The space-before value, in floating-point points applied.</returns>

        [DispId(dispId: 0x424)]
        float SpaceBefore { get; set; }

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
        [DispId(dispId: 0x425)]
        TomBoolean WidowControl { get; set; }

        /// <summary>
        ///     Retrieves the tab count.
        /// </summary>


        /// <remarks>
        ///     The tab count of a new instance can be nonzero, depending on the underlying text engine.
        ///     For example, Microsoft Word stories begin with no explicit tabs defined, while rich edit instances start with a
        ///     single explicit tab.
        ///     To be sure there are no explicit tabs (that is, to set the tab count to zero), call TextParagraph.ClearAllTabs.
        /// </remarks>
        [DispId(dispId: 0x426)]
        int TabCount { get; }

        /// <summary>
        ///     Adds a tab at the displacement tabPos, with type tabAlign, and leader style, tabLeader.
        /// </summary>
        /// <param name="tabPos">New tab displacement, in floating-point points.</param>
        /// <param name="tabAlign">A ParagraphTabAlign options for the tab position.</param>
        /// <param name="tabLeader">
        ///     A ParagraphAddTabLeader enum value. A leader character is the character that is used to fill
        ///     the space taken by a tab character.
        /// </param>

        [DispId(dispId: 0x427)]
        void AddTab([In] float tabPos, [In] ParagraphTabAlignment tabAlign, [In] ParagraphAddTabLeader tabLeader);

        /// <summary>
        ///     Clears all tabs, reverting to equally spaced tabs with the default tab spacing.
        /// </summary>

        [DispId(dispId: 0x428)]
        void ClearAllTabs();

        /// <summary>
        ///     Deletes a tab at a specified displacement.
        /// </summary>
        /// <param name="tabPos">Displacement, in floating-point points, at which a tab should be deleted.</param>

        [DispId(dispId: 0x429)]
        void DeleteTab([In] float tabPos);

        /// <summary>
        ///     Retrieves tab parameters (displacement, alignment, and leader style) for a specified tab.
        /// </summary>
        /// <param name="tabIndexOrConstant">
        ///     Index of tab for which to retrieve info.
        ///     It can be either a numerical index or a ParagraphTabIndex enum value.
        ///     Since tab indexes are zero-based, iTab = zero gets the first tab defined, iTab = 1 gets the second tab defined, and
        ///     so forth.
        /// </param>
        /// <param name="tabPos">
        ///     A variable that receives the tab displacement, in floating-point points.
        ///     The return value of tabPos is zero if the tab does not exist and the return value of tabPos is TomBoolean.Undefined
        ///     if there are multiple values in the associated range.
        /// </param>
        /// <param name="tabAlign">Receives the tab alignment. For more information, see TextParagraph.AddTab. </param>
        /// <param name="tabLeader">Receives the tab leader-character style. For more information, see TextParagraph.AddTab</param>

        [DispId(dispId: 0x430)]
        void GetTab([In] ParagraphTabIndex tabIndexOrConstant, out float tabPos, out ParagraphTabAlignment tabAlign, out ParagraphAddTabLeader tabLeader);
    }
}