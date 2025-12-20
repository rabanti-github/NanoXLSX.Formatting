/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using static NanoXLSX.Extensions.Color;

namespace NanoXLSX.Extensions
{
    /// <summary>
    /// Builder for creating inline text styles with a fluent API.
    /// </summary>
    public class InlineStyleBuilder
    {
        private readonly InlineStyle style = new InlineStyle();

        /// <summary>
        /// Sets the font name for the inline style.
        /// </summary>
        /// <param name="fontName">Font name</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Font(string fontName)
        {
            style.FontName = fontName;
            return this;
        }

        /// <summary>
        /// Sets the font size for the inline style.
        /// </summary>
        /// <param name="size">Font size in points</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Size(float size)
        {
            style.FontSize = size;
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
            style.Strikethrough = strikethrough;
            return this;
        }

        /// <summary>
        /// Sets the underline style for the font.
        /// </summary>
        /// <param name="style">Underline style</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Underline(UnderlineStyle style = UnderlineStyle.Single)
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
            style.Color = color;
            return this;
        }

        /// <summary>
        /// Sets the font color from an RGB hex string.
        /// </summary>
        /// <param name="rgb">RGB value (6 characters with optional, leading #)</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder ColorRgb(string rgb)
        {
            style.Color = Extensions.Color.FromRgb(rgb);
            return this;
        }

        /// <summary>
        /// Sets the font color from an ARGB hex string.
        /// </summary>
        /// <param name="argb">ARGB value (8 characters with optional, leading #)</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder ColorArgb(string argb)
        {
            style.Color = Extensions.Color.FromArgb(argb);
            return this;
        }

        /// <summary>
        /// Sets the font color from a theme color.
        /// </summary>
        /// <param name="theme">Color scheme</param>
        /// <param name="tint">Optional tint value (-1.0 to 1.0, default = 0.0)</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder ColorTheme(ColorScheme theme, double tint = 0.0)
        {
            style.Color = Extensions.Color.FromTheme(theme, tint);
            return this;
        }

        /// <summary>
        /// Sets the font color from an indexed color value.
        /// </summary>
        /// <param name="indexed">Indexed color value</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder ColorIndexed(IndexedColor indexed)
        {
            style.Color = Extensions.Color.FromIndexed(indexed);
            return this;
        }

        /// <summary>
        /// Sets the vertical alignment for the font.
        /// </summary>
        /// <param name="alignment">Alignment value</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder VerticalAlign(VerticalAlignment alignment)
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
            style.VerticalAlign = VerticalAlignment.Superscript;
            return this;
        }

        /// <summary>
        /// Sets the font to subscript.
        /// </summary>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Subscript()
        {
            style.VerticalAlign = VerticalAlignment.Subscript;
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
        /// <param name="charset">Font charset</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Charset(FontCharset charset)
        {
            style.Charset = charset;
            return this;
        }

        /// <summary>
        /// Sets the font family.
        /// </summary>
        /// <param name="family">Font family</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Family(FontFamily family)
        {
            style.Family = family;
            return this;
        }

        /// <summary>
        /// Sets the font scheme.
        /// </summary>
        /// <param name="scheme">Font scheme</param>
        /// <returns>The current builder instance</returns>
        public InlineStyleBuilder Scheme(FontScheme scheme)
        {
            style.Scheme = scheme;
            return this;
        }

        /// <summary>
        /// Builds the inline style instance.
        /// </summary>
        /// <param name="reset">If true, resets the builder after building (default = true)</param>
        /// <returns>The built inline style</returns>
        public InlineStyle Build(bool reset = true)
        {
            InlineStyle instance = style.Copy();
            if (reset)
            {
                style.Reset();
            }
            return instance;
        }
    }
}
