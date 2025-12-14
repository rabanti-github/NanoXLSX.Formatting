using System;
using System.Collections.Generic;
using System.Text;

namespace NanoXLSX.Extensions
{
    #region Enums

    public enum PhoneticType
    {
        HalfwidthKatakana,
        FullwidthKatakana,
        Hiragana,
        NoConversion
    }

    public enum PhoneticAlignment
    {
        NoControl,
        Left,
        Center,
        Distributed
    }

    public enum UnderlineStyle
    {
        Single,
        Double,
        SingleAccounting,
        DoubleAccounting,
        None
    }

    public enum VerticalAlignment
    {
        Baseline,
        Superscript,
        Subscript
    }

    public enum FontScheme
    {
        None,
        Major,
        Minor
    }

    public enum FontCharset
    {
        ANSI = 0,
        Default = 1,
        Symbols = 2,
        Macintosh = 77,
        JIS = 128,
        Hangul = 129,
        Johab = 130,
        GBK = 134,
        Big5 = 136,
        Greek = 161,
        Turkish = 162,
        Vietnamese = 163,
        Hebrew = 177,
        Arabic = 178,
        Baltic = 186,
        Russian = 204,
        Thai = 222,
        EasternEuropean = 238,
        OEM = 255
    }

    public enum FontFamily
    {
        NotApplicable = 0,
        Roman = 1,
        Swiss = 2,
        Modern = 3,
        Script = 4,
        Decorative = 5,
        Reserved1 = 6,
        Reserved2 = 7,
        Reserver3 = 8,
        Reserved4 = 9,
        Reserved5 = 10,
        Reserved6 = 11,
        Reserved7 = 12,
        Reserved8 = 13,
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

        public string FontName { get; set; }
        public FontCharset? Charset { get; set; }
        public FontFamily? Family { get; set; } = DefaultFontFamily;
        public bool? Bold { get; set; }
        public bool? Italic { get; set; }
        public bool? Strikethrough { get; set; }
        public bool? Outline { get; set; }
        public bool? Shadow { get; set; }
        public bool? Condense { get; set; }
        public bool? Extend { get; set; }
        public Color Color { get; set; }
        public float? FontSize { get; set; }  = DefaultFontSize;
        public UnderlineStyle? Underline { get; set; }
        public VerticalAlignment? VerticalAlign { get; set; }
        public FontScheme? Scheme { get; set; }

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
