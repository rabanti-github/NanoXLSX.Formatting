using System;
using System.Collections.Generic;
using System.Text;

namespace NanoXLSX.Extensions
{
    /// <summary>
    /// Represents a single text run with optional inline formatting.
    /// </summary>
    public class TextRun
    {
        public string Text { get; set; }
        public InlineStyle Style { get; set; }

        public TextRun(string text, InlineStyle style = null)
        {
            Text = text ?? throw new ArgumentNullException(nameof(text));
            Style = style;
        }

        public TextRun Copy()
        {
            return new TextRun(this.Text, this.Style?.Copy());
        }

    }
}
