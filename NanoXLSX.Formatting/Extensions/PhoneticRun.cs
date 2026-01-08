/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using NanoXLSX.Utils.Xml;

namespace NanoXLSX.Extensions
{
    /// <summary>
    /// Represents a phonetic run that provides pronunciation guidance for text.
    /// </summary>
    public class PhoneticRun
    {
        /// <summary>
        /// The phonetic text to be displayed (Ruby text,like Furigana, Pinyin or Zhuyin).
        /// </summary>
        public string Text { get; set; }
        /// <summary>
        /// The start index of the base text (character where the Ruby text starts)
        /// </summary>
        public uint StartBase { get; set; }
        /// <summary>
        /// The end index of the base text (character where the Ruby text ends)
        /// </summary>
        public uint EndBase { get; set; }

        /// <summary>
        /// Constructor to create a phonetic run.
        /// </summary>
        /// <param name="text">The phonetic text to be displayed (Ruby text,like Furigana, Pinyin or Zhuyin)</param>
        /// <param name="startBase">The start index of the base text (character where the Ruby text starts)</param>
        /// <param name="endBase">The end index of the base text (character where the Ruby text ends)</param>
        /// <exception cref="FormatException">Thrown when the text is null or empty</exception>
        public PhoneticRun(string text, uint startBase, uint endBase)
        {
            if (string.IsNullOrEmpty(text))
            {
                throw new FormatException("The text of the phonetic run cannot be null or empty");
            }
            Text = XmlUtils.SanitizeXmlValue(text);
            StartBase = startBase;
            EndBase = endBase;
        }

        /// <summary>
        /// Creates a copy of the current phonetic run.
        /// </summary>
        /// <returns>Deep copy of the phonetic run</returns>
        public PhoneticRun Copy()
        {
            return new PhoneticRun(this.Text, this.StartBase, this.EndBase);
        }
    }
}
