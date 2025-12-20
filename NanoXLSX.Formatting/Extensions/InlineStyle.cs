/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Extensions
{
    #region Enums

    /// <summary>
    /// Enumeration for phonetic text types.
    /// </summary>
    public enum PhoneticType
    {
        /// <summary> Half-width Katakana characters for phonetic text.</summary>
        HalfwidthKatakana,
        /// <summary>Full-width Katakana characters for phonetic text.</summary>
        FullwidthKatakana,
        /// <summary>Hiragana characters for phonetic text.</summary>
        Hiragana,
        /// <summary>No conversion for phonetic text.</summary>
        NoConversion
    }

    /// <summary>
    /// Enumeration for phonetic text alignment.
    /// </summary>
    public enum PhoneticAlignment
    {
        /// <summary>
        /// Each phonetic character is left justified without respect to the base text (so it is not per word).</summary>
        NoControl,
        /// <summary>Each phonetic character is left justified with respect to the base text., per word.</summary>
        Left,
        /// <summary> Center the phonetic characters over the base word, per word. </summary>
        Center,
        /// <summary> Each phonetic character is distributed above each base word character, per word.</summary>
        Distributed
    }

    /// <summary>
    /// Enumeration for underline styles.
    /// </summary>
    public enum UnderlineStyle
    {
        /// <summary>Text contains a single underline</summary>
#pragma warning disable CA1720 // Suppress: Identifiers should not contain type
        Single,
        /// <summary>Text contains a double underline</summary>
        Double,
#pragma warning restore CA1707
        /// <summary>Text contains a single, accounting underline</summary>
        SingleAccounting,
        /// <summary>Text contains a double, accounting underline</summary>
        DoubleAccounting,
        /// <summary>Text contains no underline (default)</summary>
        None,
    }

    /// <summary>
    /// Enumeration for vertical text alignment.
    /// </summary>
    public enum VerticalAlignment
    {
        /// <summary>Text will be rendered at the normal baseline</summary>
        Baseline,
        /// <summary>Text will be rendered as subscript</summary>
        Superscript,
        /// <summary>Text will be rendered as superscript</summary>
        Subscript
    }

    /// <summary>
    /// Enumeration for font schemes.
    /// </summary>
    public enum FontScheme
    {
        /// <summary>Font scheme is major</summary>
        Major,
        /// <summary>Font scheme is minor (default)</summary>
        Minor,
        /// <summary>No Font scheme is used</summary>
        None,
    }

    /// <summary>
    /// Enumeration for font character sets.
    /// </summary>
    public enum FontCharset
    {
        /// <summary>
        /// Application-defined (any other value than the defined enum values; can be ignored)
        /// </summary>
        ApplicationDefined = -1,
        /// <summary>
        /// Charset according to iso-8859-1
        /// </summary>
        ANSI = 0,
        /// <summary>
        /// Default charset (not defined more specific)
        /// </summary>
        Default = 1,
        /// <summary>
        /// Symbols from the private Unicode range U+FF00 to U+FFFF, to display special characters in the range of U+0000 to U+00FF
        /// </summary>
        Symbols = 2,
        /// <summary>
        /// Macintosh charset, Standard Roman
        /// </summary>
        Macintosh = 77,
        /// <summary>
        /// Shift JIS charset (shift_jis)
        /// </summary>
        JIS = 128,
        /// <summary>
        /// Hangul charset (ks_c_5601-1987)
        /// </summary>
        Hangul = 129,
        /// <summary>
        /// Johab charset (KSC-5601-1992)
        /// </summary>
        Johab = 130,
        /// <summary>
        /// GBK charset (GB-2312)
        /// </summary>
        GBK = 134,
        /// <summary>
        /// Chinese Big Five charset
        /// </summary>
        Big5 = 136,
        /// <summary>
        /// Greek charset (windows-1253)
        /// </summary>
        Greek = 161,
        /// <summary>
        /// Turkish charset (iso-8859-9)
        /// </summary>
        Turkish = 162,
        /// <summary>
        /// Vietnamese charset (windows-1258)
        /// </summary>
        Vietnamese = 163,
        /// <summary>
        /// Hebrew charset (windows-1255)
        /// </summary>
        Hebrew = 177,
        /// <summary>
        /// Arabic charset (windows-1256)
        /// </summary>
        Arabic = 178,
        /// <summary>
        /// Baltic charset (windows-1257)
        /// </summary>
        Baltic = 186,
        /// <summary>
        /// Russian charset (windows-1251)
        /// </summary>
        Russian = 204,
        /// <summary>
        /// Thai charset (windows-874)
        /// </summary>
        Thai = 222,
        /// <summary>
        /// Eastern Europe charset (windows-1250)
        /// </summary>
        EasternEuropean = 238,
        /// <summary>
        /// OEM characters, not defined by ECMA-376
        /// </summary>
        OEM = 255
    }

    /// <summary>
    /// Enumeration for font families.
    /// </summary>
    public enum FontFamily
    {
        /// <summary>
        /// The family is not defined or not applicable
        /// </summary>
        NotApplicable = 0,
        /// <summary>
        /// The specified font implements a Roman font
        /// </summary>
        Roman = 1,
        /// <summary>
        /// The specified font implements a Swiss font
        /// </summary>
        Swiss = 2,
        /// <summary>
        /// The specified font implements a Modern font
        /// </summary>
        Modern = 3,
        /// <summary>
        /// The specified font implements a Script font
        /// </summary>
        Script = 4,
        /// <summary>
        /// The specified font implements a Decorative font
        /// </summary>
        Decorative = 5,
        /// <summary>
        /// The specified font implements a not yet defined font archetype (reserved / do not use)
        /// </summary>
        Reserved1 = 6,
        /// <summary>
        /// The specified font implements a not yet defined font archetype (reserved / do not use)
        /// </summary>
        Reserved2 = 7,
        /// <summary>
        /// The specified font implements a not yet defined font archetype (reserved / do not use)
        /// </summary>
        Reserved3 = 8,
        /// <summary>
        /// The specified font implements a not yet defined font archetype (reserved / do not use)
        /// </summary>
        Reserved4 = 9,
        /// <summary>
        /// The specified font implements a not yet defined font archetype (reserved / do not use)
        /// </summary>
        Reserved5 = 10,
        /// <summary>
        /// The specified font implements a not yet defined font archetype (reserved / do not use)
        /// </summary>
        Reserved6 = 11,
        /// <summary>
        /// The specified font implements a not yet defined font archetype (reserved / do not use)
        /// </summary>
        Reserved7 = 12,
        /// <summary>
        /// The specified font implements a not yet defined font archetype (reserved / do not use)
        /// </summary>
        Reserved8 = 13,
        /// <summary>
        /// The specified font implements a not yet defined font archetype (reserved / do not use)
        /// </summary>
        Reserved9 = 14,
    }

    #endregion

    /// <summary>
    /// Represents inline text formatting properties for a text run.
    /// </summary>
    public class InlineStyle
    {

        #region Constants
        /// <summary>
        /// Default font size if not specified
        /// </summary>
        public const float DefaultFontSize = 11f;
        public const FontFamily DefaultFontFamily = FontFamily.Swiss;
        #endregion

        /// <summary>
        /// Name of the font
        /// </summary>
        public string FontName { get; set; }
        /// <summary>
        /// Character set of the font
        /// </summary>
        public FontCharset? Charset { get; set; }
        /// <summary>
        /// Font family
        /// </summary>
        public FontFamily? Family { get; set; } = DefaultFontFamily;
        /// <summary>
        /// Indicating whether the font is bold
        /// </summary>
        public bool? Bold { get; set; }
        /// <summary>
        /// Indicating whether the font is italic
        /// </summary>
        public bool? Italic { get; set; }
        /// <summary>
        /// Indicating whether the font has a strikethrough
        /// </summary>
        public bool? Strikethrough { get; set; }
        /// <summary>
        /// Indicating whether the font has an outline
        /// </summary>
        public bool? Outline { get; set; }
        /// <summary>
        /// Indicating whether the font has a shadow
        /// </summary>
        public bool? Shadow { get; set; }
        /// <summary>
        /// Indicating whether the font is condensed
        /// </summary>
        public bool? Condense { get; set; }
        /// <summary>
        /// Indicating whether the font is extended
        /// </summary>
        public bool? Extend { get; set; }
        /// <summary>
        /// Color of the font
        /// </summary>
        public Color Color { get; set; }
        /// <summary>
        /// Size of the font in points (Optional; default 11)
        /// </summary>
        public float? FontSize { get; set; } = DefaultFontSize;
        /// <summary>
        /// Underline style of the font
        /// </summary>
        public UnderlineStyle? Underline { get; set; }
        /// <summary>
        /// Vertical alignment of the font
        /// </summary>
        public VerticalAlignment? VerticalAlign { get; set; }
        /// <summary>
        /// Font scheme
        /// </summary>
        public FontScheme? Scheme { get; set; }

        /// <summary>
        /// Copies the current InlineStyle instance.
        /// </summary>
        /// <returns>InlineStyle instance</returns>
        public InlineStyle Copy()
        {
            return new InlineStyle
            {
                FontName = FontName,
                Charset = Charset,
                Family = Family,
                Bold = Bold,
                Italic = Italic,
                Strikethrough = Strikethrough,
                Outline = Outline,
                Shadow = Shadow,
                Condense = Condense,
                Extend = Extend,
                Color = Color.Copy(),
                FontSize = FontSize,
                Underline = Underline,
                VerticalAlign = VerticalAlign,
                Scheme = Scheme
            };
        }

        /// <summary>
        /// Resets all properties to their default values.
        /// </summary>
        public void Reset()
        {
            FontName = null;
            Charset = null;
            Family = DefaultFontFamily;
            Bold = null;
            Italic = null;
            Strikethrough = null;
            Outline = null;
            Shadow = null;
            Condense = null;
            Extend = null;
            Color = null;
            FontSize = DefaultFontSize;
            Underline = null;
            VerticalAlign = null;
            Scheme = null;
        }

    }
}
