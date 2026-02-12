/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Colors;
using NanoXLSX.Styles;
using NanoXLSX.Themes;

namespace NanoXLSX.Extensions
{
    /// <summary>
    /// Builder for creating inline <see cref="Styles.Font"/> styles with a fluent API.
    /// </summary>
    public class InlineStyleBuilder
    {
        private Font style = new Font();

        /// <summary>
        /// Sets the font name for the inline style.
        /// </summary>
        /// <param name="fontName">Font name</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder FontName(string fontName)
        {
            style.Name = fontName;
            return this;
        }

        /// <summary>
        /// Sets the font size for the inline style.
        /// </summary>
        /// <param name="size">Font size in points</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Size(float size)
        {
            style.Size = size;
            return this;
        }

        /// <summary>
        /// Sets whether the font is bold.
        /// </summary>
        /// <param name="bold">Bold if true (default)</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Bold(bool bold = true)
        {
            style.Bold = bold;
            return this;
        }

        /// <summary>
        /// Sets whether the font is italic.
        /// </summary>
        /// <param name="italic">Italic if true (default)</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Italic(bool italic = true)
        {
            style.Italic = italic;
            return this;
        }

        /// <summary>
        /// Sets whether the font has strikethrough.
        /// </summary>
        /// <param name="strikethrough">Has strikethrough if true (default)</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Strikethrough(bool strikethrough = true)
        {
            style.Strike = strikethrough;
            return this;
        }

        /// <summary>
        /// Sets the underline style for the font.
        /// </summary>
        /// <param name="style">Underline style</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Underline(Font.UnderlineValue style = Styles.Font.UnderlineValue.Single)
        {
            this.style.Underline = style;
            return this;
        }

        /// <summary>
        /// Sets the font color.
        /// </summary>
        /// <param name="color">Font color</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Color(Color color)
        {
            style.ColorValue = color;
            return this;
        }

        /// <summary>
        /// Sets the font color from an RGB or ARGB hex string.
        /// </summary>
        /// <param name="argb">RGB value (6 characters) or ARGB (8 characters) with optional, leading #</param>
        /// <returns>The current builder instance</returns>
        /// \remark <remarks>In case of a 6 character RGB value, 'FF' will be automatically prepended to this value</remarks>
        public InlineStyleBuilder ColorArgb(string argb)
        {
            style.ColorValue = Colors.Color.CreateRgb(argb);
            return this;
        }

        /// <summary>
        /// Sets the font color from a theme color.
        /// </summary>
        /// <param name="theme">Color scheme element</param>
        /// <param name="tint">Optional tint value (-1.0 to 1.0, default = 0.0)</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder ColorTheme(Theme.ColorSchemeElement theme, double tint = 0.0)
        {
            style.ColorValue = Colors.Color.CreateTheme(theme, tint);
            return this;
        }

        /// <summary>
        /// Sets the font color from an indexed color value.
        /// </summary>
        /// <param name="indexed">Indexed color value</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder ColorIndexed(IndexedColor.Value indexed)
        {
            style.ColorValue = Colors.Color.CreateIndexed(indexed);
            return this;
        }

        /// <summary>
        /// Sets the font color from an indexed color value, using the index number (0-65).
        /// </summary>
        /// <param name="colorIndex">Indexed color value</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder ColorIndexed(int colorIndex)
        {
            style.ColorValue = Colors.Color.CreateIndexed(colorIndex);
            return this;
        }

        /// <summary>
        /// Sets the vertical alignment for the font.
        /// </summary>
        /// <param name="alignment">Alignment value</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder VerticalAlign(Font.VerticalTextAlignValue alignment)
        {
            style.VerticalAlign = alignment;
            return this;
        }

        /// <summary>
        /// Sets the font to superscript.
        /// </summary>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Superscript()
        {
            style.VerticalAlign = Font.VerticalTextAlignValue.Superscript;
            return this;
        }

        /// <summary>
        /// Sets the font to subscript.
        /// </summary>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Subscript()
        {
            style.VerticalAlign = Font.VerticalTextAlignValue.Subscript;
            return this;
        }

        /// <summary>
        /// Sets whether the font has outline.
        /// </summary>
        /// <param name="outline">Has an outline if true (default)</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Outline(bool outline = true)
        {
            style.Outline = outline;
            return this;
        }

        /// <summary>
        /// Sets whether the font has shadow.
        /// </summary>
        /// <param name="shadow">Has a shadow if true (default)</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Shadow(bool shadow = true)
        {
            style.Shadow = shadow;
            return this;
        }

        /// <summary>
        /// Sets whether the font is condensed.
        /// </summary>
        /// <param name="condense">Is condensed if true (default)</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Condense(bool condense = true)
        {
            style.Condense = condense;
            return this;
        }

        /// <summary>
        /// Sets whether the font is extended.
        /// </summary>
        /// <param name="extend">Is extended if true (default)</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Extend(bool extend = true)
        {
            style.Extend = extend;
            return this;
        }

        /// <summary>
        /// Sets the font charset.
        /// </summary>
        /// <param name="charset">Font charset value</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Charset(Font.CharsetValue charset)
        {
            style.Charset = charset;
            return this;
        }

        /// <summary>
        /// Sets the font family.
        /// </summary>
        /// <param name="family">Font family value</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Family(Font.FontFamilyValue family)
        {
            style.Family = family;
            return this;
        }

        /// <summary>
        /// Sets the font scheme.
        /// </summary>
        /// <param name="scheme">Font scheme value</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Scheme(Font.SchemeValue scheme)
        {
            style.Scheme = scheme;
            return this;
        }

        /// <summary>
        /// Builds the inline style instance.
        /// </summary>
        /// <param name="reset">If true, resets the builder after building (default = true)</param>
        /// <returns>The built inline style</returns>
        public Font Build(bool reset = true)
        {
            Font instance = style.CopyFont();
            if (reset)
            {
                style = new Font();
            }
            return instance;
        }
    }
}
