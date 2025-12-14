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
        private readonly InlineStyle _style = new InlineStyle();

        public InlineStyleBuilder Font(string fontName)
        {
            _style.FontName = fontName;
            return this;
        }

        public InlineStyleBuilder Size(double size)
        {
            _style.FontSize = size;
            return this;
        }

        public InlineStyleBuilder Bold(bool bold = true)
        {
            _style.Bold = bold;
            return this;
        }

        public InlineStyleBuilder Italic(bool italic = true)
        {
            _style.Italic = italic;
            return this;
        }

        public InlineStyleBuilder Strikethrough(bool strikethrough = true)
        {
            _style.Strikethrough = strikethrough;
            return this;
        }

        public InlineStyleBuilder Underline(UnderlineStyle style = UnderlineStyle.Single)
        {
            _style.Underline = style;
            return this;
        }

        public InlineStyleBuilder Color(Color color)
        {
            _style.Color = color;
            return this;
        }

        public InlineStyleBuilder ColorRgb(string rgb)
        {
            _style.Color = NanoXLSX.Extensions.Color.FromRgb(rgb);
            return this;
        }

        public InlineStyleBuilder ColorTheme(uint theme, double tint = 0.0)
        {
            _style.Color = NanoXLSX.Extensions.Color.FromTheme(theme, tint);
            return this;
        }

        public InlineStyleBuilder ColorIndexed(uint indexed)
        {
            _style.Color = NanoXLSX.Extensions.Color.FromIndexed(indexed);
            return this;
        }

        public InlineStyleBuilder VerticalAlign(VerticalAlignment alignment)
        {
            _style.VerticalAlign = alignment;
            return this;
        }

        public InlineStyleBuilder Superscript()
        {
            _style.VerticalAlign = VerticalAlignment.Superscript;
            return this;
        }

        public InlineStyleBuilder Subscript()
        {
            _style.VerticalAlign = VerticalAlignment.Subscript;
            return this;
        }

        public InlineStyleBuilder Outline(bool outline = true)
        {
            _style.Outline = outline;
            return this;
        }

        public InlineStyleBuilder Shadow(bool shadow = true)
        {
            _style.Shadow = shadow;
            return this;
        }

        public InlineStyleBuilder Condense(bool condense = true)
        {
            _style.Condense = condense;
            return this;
        }

        public InlineStyleBuilder Extend(bool extend = true)
        {
            _style.Extend = extend;
            return this;
        }

        public InlineStyleBuilder Charset(FontCharset charset)
        {
            _style.Charset = charset;
            return this;
        }

        public InlineStyleBuilder Family(FontFamily family)
        {
            _style.Family = family;
            return this;
        }

        public InlineStyleBuilder Scheme(FontScheme scheme)
        {
            _style.Scheme = scheme;
            return this;
        }

        public InlineStyle Build()
        {
            return _style;
        }
    }
}
