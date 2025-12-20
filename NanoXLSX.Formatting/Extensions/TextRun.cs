/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;

namespace NanoXLSX.Extensions
{
    /// <summary>
    /// Represents a single text run with optional inline formatting.
    /// </summary>
    public class TextRun
    {
        /// <summary>
        /// Plain text of the run.
        /// </summary>
        public string Text { get; set; }
        /// <summary>
        /// Inline style applied to the text run.
        /// </summary>
        public InlineStyle Style { get; set; }

        /// <summary>
        /// Constructor to create a text run with optional inline style.
        /// </summary>
        /// <param name="text">Plain text</param>
        /// <param name="style">Optional inline style</param>
        /// <exception cref="Exceptions.FormatException">Thrown when the text is null</exception>
        public TextRun(string text, InlineStyle style = null)
        {
            Text = text ?? throw new Exceptions.FormatException("The text of a text run cannot be null");
            Style = style;
        }

        /// <summary>
        /// Creates a copy of the current text run.
        /// </summary>
        /// <returns>Copy of the text run</returns>
        public TextRun Copy()
        {
            return new TextRun(this.Text, this.Style?.Copy());
        }

    }
}
