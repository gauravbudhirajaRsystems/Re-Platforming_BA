// © Copyright 2018 Levit & James, Inc.

using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace LevitJames.TextServices
{
    /// <summary>
    ///     Represents the Text Object Model interface ITextFont.
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
    ///         TextFont uses a special variable enum type, called TomBoolean, for rich-text attributes that have binary
    ///         states.
    ///         The TomBoolean enum is different from the Boolean type because it can take four values: TomBoolean.True,
    ///         TomBoolean.False, TomBoolean.Toggle, and TomBoolean.Undefined.
    ///         The TomBoolean.True and tomFalse values indicate True and False. The TomBoolean.Toggle value is used to toggle
    ///         a property.
    ///         The TomBoolean.Undefined value, more traditionally called NINCH, is a special no-input,
    ///         no-change value that works with longs, floats, and COLOREFs. For strings, TomBoolean.Undefined (or NINCH) is
    ///         represented by the null string. For Set operations, using tomUndefined does not change the target property.
    ///         For Get operations, TomBoolean.Undefined means that the characters in the range have different values (it gives
    ///         the grayed check box in property dialog boxes)
    ///     </para>
    ///     <para>
    ///         The rich edit control is able to accept and return all TextFont properties intact, that is, without
    ///         modification, both through TOM and through its
    ///         Rich Text Format (RTF) converters. However, it cannot display the All Caps, Animation, Embossed, Imprint,
    ///         Shadow, Small Caps, Hidden, Kerning, Outline,
    ///         and Style font Properties.Resources.
    ///     </para>
    ///     <para>
    ///         This is a Com interface so Marshal.ReleaseComObject should be used to free the object when you have finished
    ///         with it.
    ///     </para>
    /// </remarks>
    [CompilerGenerated]
    [SuppressMessage("Microsoft.Naming", "CA1715:IdentifiersShouldHaveCorrectPrefix", MessageId = "I")]
    [ComImport]
    [TypeLibType(flags: 0x10C0)]
    [Guid("8CC497C3-A1DF-11CE-8098-00AA0047BE5D")]
    [DefaultMember("Duplicate")]
    internal interface ITextFont
    {
        /// <summary>
        ///     Creates a duplicate of this character format object.
        /// </summary>

        /// <returns>Returns a new TextFont instance.</returns>

        [DispId(dispId: 0)]
        ITextFont Duplicate { get; set; }

        /// <summary>
        ///     Determines whether the font can be changed.
        /// </summary>

        /// <remarks>
        ///     The CanChange returns True only if the font can be changed.
        ///     That is, no part of an associated range is protected and an associated document is not read-only.
        ///     If this TextFont object is a duplicate, no protection rules apply.
        /// </remarks>
        [return: MarshalAs(UnmanagedType.I4)]
        [DispId(dispId: 0x301)]
        TomBoolean CanChange();

        /// <summary>
        ///     Determines whether a specified TextFont object has the same properties as this TextFont object.
        /// </summary>
        /// <param name="font">A TextFont object that is compared to this TextFont object.</param>
        /// <returns>True if the Font objects are the same, False otherwise.</returns>
        /// <remarks>
        ///     The font objects are equal only if <paramref>font</paramref> belongs to the same Text Object Model (TOM) object as
        ///     the current font object.
        ///     The IsEqual method ignores entries for which either font object has a TomBoolean.Undefined value.
        /// </remarks>
        [return: MarshalAs(UnmanagedType.I4)]
        [DispId(dispId: 770)]
        TomBoolean IsEqual([In] [MarshalAs(UnmanagedType.Interface)]
                           ITextFont font);

        /// <summary>
        ///     Resets the character formatting to the default values.
        /// </summary>
        /// <param name="value">Default character formatting values.</param>
        /// <remarks>
        ///     Calling Reset with ResetTextFontValue.Undefined sets all properties to undefined values.
        ///     Thus, applying the font object to a range changes nothing. This applies to a font object that is obtained by the
        ///     Duplicate method.
        /// </remarks>
        [DispId(dispId: 0x303)]
        void Reset([In] ResetTextFontValue value);

        /// <summary>
        ///     Gets or sets the character style handle for the characters in a range.
        /// </summary>
        /// <value>New character style handle.</value>

        /// <remarks>
        ///     The Text Object Model (TOM) version 1.0 has no way to specify the meanings of the style handles,
        ///     which depend on other facilities of the text system implementing TOM.
        /// </remarks>
        [DispId(dispId: 0x304)]
        int Style { get; set; }

        /// <summary>
        ///     Gets or sets the state of the AllCaps property.
        /// </summary>
        /// <value>True when the AllCaps property is on and receives False when the AllCaps property is off.</value>


        [DispId(dispId: 0x305)]
        TomBoolean AllCaps { get; set; }

        /// <summary>
        ///     Gets or sets the animation type.
        /// </summary>
        /// <value>One of the values specified in the FontAnimation enum.</value>


        [DispId(dispId: 0x306)]
        FontAnimation Animation { get; set; }

        /// <summary>
        ///     Gets or sets the background color.
        /// </summary>
        /// <value>A values to Win32 Color value indicate the background (highlight) color for a range.</value>

        /// <remarks>
        ///     In .Net you can convert between Win32 COLORREF values and .Net values by using the
        ///     System.Drawing.ColorTranslator.FromOle() and System.Drawing.ColorTranslator.ToOle() methods respectively.
        /// </remarks>
        /// The property can also return or be set to a value of (-9999997) which indicates the range uses the default system background color.
        [DispId(dispId: 0x307)]
        int BackColor { get; set; }

        /// <summary>
        ///     Gets or sets the bold state.
        /// </summary>

        /// <returns>True if the Bold property is on or False if it is off.</returns>
        /// <remarks>
        ///     You can use the TextFont.Weight property to set or retrieve the font weight more precisely than the binary
        ///     Bold property.
        /// </remarks>
        [DispId(dispId: 0x308)]
        TomBoolean Bold { get; set; }

        /// <summary>
        ///     Gets or sets the state of the Emboss property, which indicates whether characters should be embossed.
        /// </summary>

        /// <returns>True if the Emboss property is on, or False if it is off.</returns>

        [DispId(dispId: 0x309)]
        TomBoolean Emboss { get; set; }

        /// <summary>
        ///     Gets or sets the foreground (text) color.
        /// </summary>
        /// <value>A values to Win32 Color value indicate the foreground (text) color for a range.</value>

        /// <remarks>
        ///     In .Net you can convert between Win32 COLORREF values and .Net values by using the
        ///     System.Drawing.ColorTranslator.FromOle() and System.Drawing.ColorTranslator.ToOle() methods respectively.
        /// </remarks>
        /// The property can also return or be set to a value of (-9999997) which indicates the range uses the default system background color.
        [DispId(dispId: 0x310)]
        int ForeColor { get; set; }

        /// <summary>
        ///     Gets or sets the state of the Hidden property, which indicates whether characters are displayed.
        /// </summary>

        /// <returns>True if the Hidden property is on or False if it is off.</returns>

        [DispId(dispId: 0x311)]
        TomBoolean Hidden { get; set; }

        /// <summary>
        ///     Gets or sets the state of the Engrave property, which indicates whether characters should be displayed as imprinted
        ///     characters.
        /// </summary>

        /// <returns>True if the Engrave property is on or False if it is off.</returns>

        [DispId(dispId: 0x312)]
        TomBoolean Engrave { get; set; }

        /// <summary>
        ///     Gets or sets the italic state.
        /// </summary>
        /// <value>True if the Italic property is on or False if it is off.</value>


        [DispId(dispId: 0x313)]
        TomBoolean Italic { get; set; }

        /// <summary>
        ///     Gets or sets the minimum kerning size, which is given in floating-point points.
        /// </summary>

        /// <returns>The font size above which kerning is turned on, in floating-point points.</returns>
        /// <remarks>
        ///     If the value pointed to by pValue is zero, kerning is turned off. Positive values turn on pair kerning for font
        ///     point sizes greater than or equal to the kerning value. For example,
        ///     the value 1 turns on kerning for all legible sizes, whereas 16 turns on kerning only for font sizes of 16 points
        ///     and larger.
        /// </remarks>
        [DispId(dispId: 0x314)]
        float Kerning { get; set; }

        /// <summary>
        ///     Gets or sets the language identifier (more generally the LCID).
        /// </summary>


        /// <remarks>
        ///     The low word of pValue contains the language identifier.
        ///     The high word is either zero or it contains the high word of the locale identifier (LCID).
        ///     To retrieve the language identifier, mask out the high word. For more information, see Locale Identifiers.
        /// </remarks>
        [DispId(dispId: 0x315)]
        int LanguageID { get; set; }

        /// <summary>
        ///     Gets or sets the font name.
        /// </summary>
        /// <value>A System.String that represents the name of the font.</value>


        [DispId(dispId: 790)]
        string Name { get; set; }

        /// <summary>
        ///     Gets or sets the state of the Outline property, which indicates whether characters are displayed as outlined
        ///     characters.
        /// </summary>

        /// <returns>True if the Outline property is on or False if it is off. </returns>
        /// <remarks>Characters are displayed as outlined characters. The value does not affect how the control displays the text.</remarks>
        [DispId(dispId: 0x317)]
        TomBoolean Outline { get; set; }

        /// <summary>
        ///     Gets or sets the character offset relative to the baseline. The value is given in floating-point points.
        /// </summary>
        /// <value>The relative vertical offset, in floating-point points.</value>

        /// <remarks>
        ///     Normally, displayed text has a zero value for this property. Positive values raise the text, and negative
        ///     values lower it.
        /// </remarks>
        [DispId(dispId: 0x318)]
        float Position { get; set; }

        /// <summary>
        ///     Gets or sets the state of the Protected property, which indicates whether the characters are protected against
        ///     modification attempts.
        /// </summary>



        [DispId(dispId: 0x319)]
        TomBoolean Protected { get; set; }

        /// <summary>
        ///     Gets or sets the state of the Shadow property, which indicates whether characters should be displayed as shadowed
        ///     characters.
        /// </summary>
        /// <value>True if the Shadow property is on or False if it is off. </value>


        [DispId(dispId: 800)]
        TomBoolean Shadow { get; set; }

        /// <summary>
        ///     Gets or sets the font size in floating-point points.
        /// </summary>
        /// <value>New font size, in floating-point points.</value>
        /// <returns>The current font size, in floating-point points</returns>

        [DispId(dispId: 0x321)]
        float Size { get; set; }

        /// <summary>
        ///     Gets or sets the state of the SmallCaps property.
        /// </summary>

        /// <returns>True if the SmallCaps property is on or False if it is off.</returns>

        [DispId(dispId: 0x322)]
        TomBoolean SmallCaps { get; set; }

        /// <summary>
        ///     Determines the inter-character spacing.
        /// </summary>
        /// <value>The horizontal spacing between characters, in floating-point points.</value>

        /// <remarks>
        ///     Normally, spaced text is given by an inter-character spacing value of zero. Positive values expand the
        ///     spacing, and negative values compress it.
        /// </remarks>
        [DispId(dispId: 0x323)]
        float Spacing { get; set; }

        /// <summary>
        ///     Gets or sets the state of the StrikeThrough property, which indicates whether characters should be displayed as
        ///     struck out.
        /// </summary>

        /// <returns>True if the StrikeThrough property is on or False if it is off.</returns>

        [DispId(dispId: 0x324)]
        TomBoolean StrikeThrough { get; set; }

        /// <summary>
        ///     Gets or sets the state of the Subscript property, which indicates whether characters are displayed as subscript.
        /// </summary>

        /// <returns>True if the Subscript property is on or False if it is off.</returns>

        [DispId(dispId: 0x325)]
        TomBoolean Subscript { get; set; }

        /// <summary>
        ///     Gets or sets the state of the Superscript property.
        /// </summary>
        /// <value>True if the Superscript property is on or False if it is off.</value>


        [DispId(dispId: 0x326)]
        TomBoolean Superscript { get; set; }

        /// <summary>
        ///     Gets or sets the Underline style.
        /// </summary>

        /// <returns>A FontUnderline Enum value representing the underline style.</returns>

        [DispId(dispId: 0x327)]
        FontUnderline Underline { get; set; }

        /// <summary>
        ///     Gets or sets the font weight for the characters in a range.
        /// </summary>

        /// <returns>A FontWeight Enum value representing the weight of the font.</returns>
        /// <remarks>The Bold property is a binary version of the Weight property that sets the weight to FontWeight.Bold</remarks>
        [DispId(dispId: 0x328)]
        FontWeight Weight { get; set; }
    }
}