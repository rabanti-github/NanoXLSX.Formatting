/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Styles;
using NanoXLSX.Utils.Xml;

namespace NanoXLSX.Extensions
{
    /// <summary>
    /// Represents a single text run with optional inline formatting.
    /// </summary>
    public class TextRun
    {
        private string text;

        /// <summary>
        /// Plain text of the run.
        /// </summary>
        /// <exception cref="Exceptions.FormatException">Thrown when the text is null</exception>
        public string Text 
        {
            get { return text; }
            set
            {
                if (value == null)
                {
                    throw new Exceptions.FormatException("The text of a text run cannot be null");
                }
                text = XmlUtils.SanitizeXmlValue(value); ;
            }
        }
        /// <summary>
        /// Font style applied to the text run.
        /// </summary>
        public Font FontStyle { get; set; }

        /// <summary>
        /// Constructor to create a text run with optional inline style.
        /// </summary>
        /// <param name="text">Plain text</param>
        /// <param name="fontStyle">Optional font style</param>
        /// <exception cref="Exceptions.FormatException">Thrown when the text is null</exception>
        public TextRun(string text, Font fontStyle = null)
        {
            Text = text;
            FontStyle = fontStyle;
        }

        /// <summary>
        /// Creates a copy of the current text run.
        /// </summary>
        /// <returns>Copy of the text run</returns>
        public TextRun Copy()
        {
            return new TextRun(this.Text, this.FontStyle?.CopyFont());
        }

    }
}
