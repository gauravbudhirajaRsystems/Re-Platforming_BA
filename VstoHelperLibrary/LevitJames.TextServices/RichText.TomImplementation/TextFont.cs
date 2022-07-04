// © Copyright 2018 Levit & James, Inc.

// ReSharper disable once RedundantUsingDirective
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using LevitJames.Core; // used with TRACK_DISPOSED

namespace LevitJames.TextServices
{
#pragma warning disable CS0659 // Type overrides Object.Equals(object o) but does not override Object.GetHashCode()


    /// <summary>
    /// Contains the Font properties of a TextRange object. Wraps the Com ITextFont interface
    /// </summary>
    /// <remarks>Open a web browser browser and search for ITextFont.xxxx for detailed documentation.</remarks>
    //[DebuggerStepThrough]
    public sealed class TextFont : ITextFont, IDisposable, IEquatable<TextFont>
#pragma warning restore CS0659 // Type overrides Object.Equals(object o) but does not override Object.GetHashCode()
    {

        private const int tomLink = -2147483616;

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private ITextFont _textFont;
#if (TRACK_DISPOSED)
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private readonly string _disposedSource;
#endif

        internal TextFont(ITextFont textFont)
        {
            _textFont = textFont;
#if (TRACK_DISPOSED)
            _disposedSource = LevitJames.Core.Diagnostics.DisposeTracker.GetSource();
#endif
        }

        //public ITextFont2 GetITextFont2() => _textFont as ITextFont2;

        public bool? LinkEffect
        {
            get
            {
                if (!(_textFont is ITextFont2 itf2))
                    return false;

                var mask = 0;
                var value = 0;
                itf2.GetEffects(out value, out mask);

                if ((mask & tomLink) != 0)
                    return ((value & tomLink) != 0);

                return null;

            }
            set
            {
                if (!(_textFont is ITextFont2 itf2))
                    return;

                itf2?.SetEffects((bool)value ? tomLink : 0, tomLink);

            }
        }

        /// <summary>
        ///     Creates a duplicate of this character format object.
        /// </summary>

        /// <returns>Returns a new TextFont instance.</returns>

        public TextFont Duplicate
        {
            get => new TextFont(_textFont.Duplicate);
            set => _textFont.Duplicate = value;
        }

        [DebuggerHidden]
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        ITextFont ITextFont.Duplicate
        {
            get => _textFont.Duplicate;
            set => _textFont.Duplicate = value;
        }


        /// <summary>
        ///     Determines whether the font can be changed.
        /// </summary>

        /// <remarks>
        ///     The CanChange returns True only if the font can be changed.
        ///     That is, no part of an associated range is protected and an associated document is not read-only.
        ///     If this TextFont object is a duplicate, no protection rules apply.
        /// </remarks>
        public bool CanChange() => _textFont.CanChange() == TomBoolean.True;

        TomBoolean ITextFont.CanChange() => _textFont.CanChange();

        TomBoolean ITextFont.IsEqual(ITextFont font) => _textFont.IsEqual(font);

        /// <summary>
        ///     Resets the character formatting to the default values.
        /// </summary>
        /// <param name="value">Default character formatting values.</param>
        /// <remarks>
        ///     Calling Reset with ResetTextFontValue.Undefined sets all properties to undefined values.
        ///     Thus, applying the font object to a range changes nothing. This applies to a font object that is obtained by the
        ///     Duplicate method.
        /// </remarks>
        public void Reset(ResetTextFontValue value)
        {
            _textFont.Reset(value);
        }

        /// <summary>
        ///     Gets or sets the character style handle for the characters in a range.
        /// </summary>
        /// <value>New character style handle.</value>

        /// <remarks>
        ///     The Text Object Model (TOM) version 1.0 has no way to specify the meanings of the style handles,
        ///     which depend on other facilities of the text system implementing TOM.
        /// </remarks>
        public int Style
        {
            get => _textFont.Style;
            set => _textFont.Style = value;
        }


        /// <summary>
        ///     Gets or sets the state of the AllCaps property.
        /// </summary>
        /// <value>True when the AllCaps property is on and receives False when the AllCaps property is off.</value>


        public TomBoolean AllCaps
        {
            get => _textFont.AllCaps;
            set => _textFont.AllCaps = value;
        }

        [DebuggerHidden]
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        TomBoolean ITextFont.AllCaps
        {
            get => _textFont.AllCaps;
            set => _textFont.AllCaps = value;
        }


        /// <summary>
        ///     Gets or sets the animation type.
        /// </summary>
        /// <value>One of the values specified in the FontAnimation enum.</value>


        public FontAnimation Animation
        {
            get => _textFont.Animation;
            set => _textFont.Animation = value;
        }

        /// <summary>
        ///     Gets or sets the background color.
        /// </summary>
        /// <value>A values to Win32 Color value indicate the background (highlight) color for a range.</value>

        /// <remarks>
        ///     In .Net you can convert between Win32 COLORREF values and .Net values by using the
        ///     System.Drawing.ColorTranslator.FromOle() and System.Drawing.ColorTranslator.ToOle() methods respectively.
        /// </remarks>
        /// The property can also return or be set to a value of (-9999997) which indicates the range uses the default system background color.
        public int BackColor
        {
            get => _textFont.BackColor;
            set => _textFont.BackColor = value;
        }

        /// <summary>
        ///     Gets or sets the bold state.
        /// </summary>

        /// <returns>True if the Bold property is on or False if it is off.</returns>
        /// <remarks>
        ///     You can use the TextFont.Weight property to set or retrieve the font weight more precisely than the binary
        ///     Bold property.
        /// </remarks>
        public TomBoolean Bold
        {
            get => _textFont.Bold;
            set => _textFont.Bold = value;
        }

        /// <summary>
        ///     Gets or sets the state of the Emboss property, which indicates whether characters should be embossed.
        /// </summary>

        /// <returns>True if the Emboss property is on, or False if it is off.</returns>

        public TomBoolean Emboss
        {
            get => _textFont.Emboss;
            set => _textFont.Emboss = value;
        }

        /// <summary>
        ///     Gets or sets the foreground (text) color.
        /// </summary>
        /// <value>A values to Win32 Color value indicate the foreground (text) color for a range.</value>

        /// <remarks>
        ///     In .Net you can convert between Win32 COLORREF values and .Net values by using the
        ///     System.Drawing.ColorTranslator.FromOle() and System.Drawing.ColorTranslator.ToOle() methods respectively.
        /// </remarks>
        /// The property can also return or be set to a value of (-9999997) which indicates the range uses the default system background color.
        public int ForeColor
        {
            get => _textFont.ForeColor;
            set => _textFont.ForeColor = value;
        }

        /// <summary>
        ///     Gets or sets the state of the Hidden property, which indicates whether characters are displayed.
        /// </summary>

        /// <returns>True if the Hidden property is on or False if it is off.</returns>

        public TomBoolean Hidden
        {
            get => _textFont.Hidden;
            set => _textFont.Hidden = value;
        }

        /// <summary>
        ///     Gets or sets the state of the Engrave property, which indicates whether characters should be displayed as imprinted
        ///     characters.
        /// </summary>

        /// <returns>True if the Engrave property is on or False if it is off.</returns>

        public TomBoolean Engrave
        {
            get => _textFont.Engrave;
            set => _textFont.Engrave = value;
        }

        /// <summary>
        ///     Gets or sets the italic state.
        /// </summary>
        /// <value>True if the Italic property is on or False if it is off.</value>


        public TomBoolean Italic
        {
            get => _textFont.Italic;
            set => _textFont.Italic = value;
        }

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
        public float Kerning
        {
            get => _textFont.Kerning;
            set => _textFont.Kerning = value;
        }

        /// <summary>
        ///     Gets or sets the language identifier (more generally the LCID).
        /// </summary>


        /// <remarks>
        ///     The low word of pValue contains the language identifier.
        ///     The high word is either zero or it contains the high word of the locale identifier (LCID).
        ///     To retrieve the language identifier, mask out the high word. For more information, see Locale Identifiers.
        /// </remarks>
        public int LanguageID
        {
            get => _textFont.LanguageID;
            set => _textFont.LanguageID = value;
        }

        /// <summary>
        ///     Gets or sets the font name.
        /// </summary>
        /// <value>A System.String that represents the name of the font.</value>


        public string Name
        {
            get => _textFont.Name;
            set => _textFont.Name = value;
        }

        /// <summary>
        ///     Gets or sets the state of the Outline property, which indicates whether characters are displayed as outlined
        ///     characters.
        /// </summary>

        /// <returns>True if the Outline property is on or False if it is off. </returns>
        /// <remarks>Characters are displayed as outlined characters. The value does not affect how the control displays the text.</remarks>
        public TomBoolean Outline
        {
            get => _textFont.Outline;
            set => _textFont.Outline = value;
        }

        /// <summary>
        ///     Gets or sets the character offset relative to the baseline. The value is given in floating-point points.
        /// </summary>
        /// <value>The relative vertical offset, in floating-point points.</value>

        /// <remarks>
        ///     Normally, displayed text has a zero value for this property. Positive values raise the text, and negative
        ///     values lower it.
        /// </remarks>
        public float Position
        {
            get => _textFont.Position;
            set => _textFont.Position = value;
        }

        /// <summary>
        ///     Gets or sets the state of the Protected property, which indicates whether the characters are protected against
        ///     modification attempts.
        /// </summary>



        public TomBoolean Protected
        {
            get => _textFont.Protected;
            set => _textFont.Protected = value;
        }

        /// <summary>
        ///     Gets or sets the state of the Shadow property, which indicates whether characters should be displayed as shadowed
        ///     characters.
        /// </summary>
        /// <value>True if the Shadow property is on or False if it is off. </value>


        public TomBoolean Shadow
        {
            get => _textFont.Shadow;
            set => _textFont.Shadow = value;
        }

        /// <summary>
        ///     Gets or sets the font size in floating-point points.
        /// </summary>
        /// <value>New font size, in floating-point points.</value>
        /// <returns>The current font size, in floating-point points</returns>

        public float Size
        {
            get => _textFont.Size;
            set => _textFont.Size = value;
        }

        /// <summary>
        ///     Gets or sets the state of the SmallCaps property.
        /// </summary>

        /// <returns>True if the SmallCaps property is on or False if it is off.</returns>

        public TomBoolean SmallCaps
        {
            get => _textFont.SmallCaps;
            set => _textFont.SmallCaps = value;
        }

        /// <summary>
        ///     Determines the inter-character spacing.
        /// </summary>
        /// <value>The horizontal spacing between characters, in floating-point points.</value>

        /// <remarks>
        ///     Normally, spaced text is given by an inter-character spacing value of zero. Positive values expand the
        ///     spacing, and negative values compress it.
        /// </remarks>
        public float Spacing
        {
            get => _textFont.Spacing;
            set => _textFont.Spacing = value;
        }

        /// <summary>
        ///     Gets or sets the state of the StrikeThrough property, which indicates whether characters should be displayed as
        ///     struck out.
        /// </summary>

        /// <returns>True if the StrikeThrough property is on or False if it is off.</returns>

        public TomBoolean StrikeThrough
        {
            get => _textFont.StrikeThrough;
            set => _textFont.StrikeThrough = value;
        }

        /// <summary>
        ///     Gets or sets the state of the Subscript property, which indicates whether characters are displayed as subscript.
        /// </summary>

        /// <returns>True if the Subscript property is on or False if it is off.</returns>

        public TomBoolean Subscript
        {
            get => _textFont.Subscript;
            set => _textFont.Subscript = value;
        }

        /// <summary>
        ///     Gets or sets the state of the Superscript property.
        /// </summary>
        /// <value>True if the Superscript property is on or False if it is off.</value>


        public TomBoolean Superscript
        {
            get => _textFont.Superscript;
            set => _textFont.Superscript = value;
        }

        /// <summary>
        ///     Gets or sets the Underline style.
        /// </summary>

        /// <returns>A FontUnderline Enum value representing the underline style.</returns>

        public FontUnderline Underline
        {
            get => _textFont.Underline;
            set => _textFont.Underline = value;
        }

        /// <summary>
        ///     Gets or sets the font weight for the characters in a range.
        /// </summary>

        /// <returns>A FontWeight Enum value representing the weight of the font.</returns>
        /// <remarks>The Bold property is a binary version of the Weight property that sets the weight to FontWeight.Bold</remarks>
        public FontWeight Weight
        {
            get => _textFont.Weight;
            set => _textFont.Weight = value;
        }



#pragma warning disable 659

        /// <summary>
        ///     Determines whether a specified TextFont object has the same properties as this TextFont object.
        /// </summary>
        /// <param name="obj">A TextFont object that is compared to this TextFont object.</param>
        /// <returns>True if the Font objects are the same, False otherwise.</returns>
        /// <remarks>
        ///     The font objects are equal only if <paramref>font</paramref> belongs to the same Text Object Model (TOM) object as
        ///     the current font object.
        ///     The IsEqual method ignores entries for which either font object has a TomBoolean.Undefined value.
        /// </remarks>
        public override bool Equals(object obj) => Equals(obj as TextFont);
#pragma warning restore 659

        /// <summary>
        ///     Determines whether a specified TextFont object has the same properties as this TextFont object.
        /// </summary>
        /// <param name="other">A TextFont object that is compared to this TextFont object.</param>
        /// <returns>True if the Font objects are the same, False otherwise.</returns>
        /// <remarks>
        ///     The font objects are equal only if <paramref>font</paramref> belongs to the same Text Object Model (TOM) object as
        ///     the current font object.
        ///     The IsEqual method ignores entries for which either font object has a TomBoolean.Undefined value.
        /// </remarks>
        public bool Equals(TextFont other)
        {
            if (other == null)
                return false;
            return _textFont.IsEqual(other._textFont) == TomBoolean.True;
        }


        // ReSharper disable once UnusedParameter.Local
        private void Dispose(bool disposeDotNetObjects)
        {
            if (_textFont != null)
            {
                Marshal.ReleaseComObject(_textFont);
                _textFont = null;
            }
        }

        /// <summary>
        /// Disposes of any resources used by the TextFont
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// The finalizer for the TextFont
        /// </summary>
        ~TextFont()
        {
#if (TRACK_DISPOSED)
            LevitJames.Core.Diagnostics.DisposeTracker.Add(_disposedSource);
#endif
            Dispose(false);
        }
    }
}