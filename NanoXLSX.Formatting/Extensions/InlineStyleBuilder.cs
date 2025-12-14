using System;
using System.Collections.Generic;
using System.Text;
using Color = NanoXLSX.Extensions.Color;

namespace NanoXLSX.Extensions
{
    /// <summary>
    /// Builder for creating inline text styles with a fluent API.
    /// </summary>
    public class InlineStyleBuilder
    {
        private readonly InlineStyle style = new InlineStyle();

        public InlineStyleBuilder Font(string fontName)
        {
            style.FontName = fontName;
            return this;
        }

        public InlineStyleBuilder Size(float size)
        {
            style.FontSize = size;
            return this;
        }

        public InlineStyleBuilder Bold(bool bold = true)
        {
            style.Bold = bold;
            return this;
        }

        public InlineStyleBuilder Italic(bool italic = true)
        {
            style.Italic = italic;
            return this;
        }

        public InlineStyleBuilder Strikethrough(bool strikethrough = true)
        {
            style.Strikethrough = strikethrough;
            return this;
        }

        public InlineStyleBuilder Underline(UnderlineStyle style = UnderlineStyle.Single)
        {
            this.style.Underline = style;
            return this;
        }

        public InlineStyleBuilder Color(Color color)
        {
            style.Color = color;
            return this;
        }

        public InlineStyleBuilder ColorRgb(string rgb)
        {
            style.Color = Extensions.Color.FromRgb(rgb);
            return this;
        }

        public InlineStyleBuilder ColorTheme(uint theme, double tint = 0.0)
        {
            style.Color = Extensions.Color.FromTheme(theme, tint);
            return this;
        }

        public InlineStyleBuilder ColorIndexed(uint indexed)
        {
            style.Color = Extensions.Color.FromIndexed(indexed);
            return this;
        }

        public InlineStyleBuilder VerticalAlign(VerticalAlignment alignment)
        {
            style.VerticalAlign = alignment;
            return this;
        }

        public InlineStyleBuilder Superscript()
        {
            style.VerticalAlign = VerticalAlignment.Superscript;
            return this;
        }

        public InlineStyleBuilder Subscript()
        {
            style.VerticalAlign = VerticalAlignment.Subscript;
            return this;
        }

        public InlineStyleBuilder Outline(bool outline = true)
        {
            style.Outline = outline;
            return this;
        }

        public InlineStyleBuilder Shadow(bool shadow = true)
        {
            style.Shadow = shadow;
            return this;
        }

        public InlineStyleBuilder Condense(bool condense = true)
        {
            style.Condense = condense;
            return this;
        }

        public InlineStyleBuilder Extend(bool extend = true)
        {
            style.Extend = extend;
            return this;
        }

        public InlineStyleBuilder Charset(FontCharset charset)
        {
            style.Charset = charset;
            return this;
        }

        public InlineStyleBuilder Family(FontFamily family)
        {
            style.Family = family;
            return this;
        }

        public InlineStyleBuilder Scheme(FontScheme scheme)
        {
            style.Scheme = scheme;
            return this;
        }

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
