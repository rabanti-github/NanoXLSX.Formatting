/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Exceptions;
using NanoXLSX.Utils.Xml;

namespace NanoXLSX.Extensions
{
    /// <summary>
    /// Represents a phonetic run that provides pronunciation guidance for text.
    /// </summary>
    public class PhoneticRun
    {

        #region enums
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
        #endregion

        private string text;

        /// <summary>
        /// The phonetic text to be displayed (Ruby text,like Furigana, Pinyin or Zhuyin).
        /// </summary>
        /// <exception cref="FormatException">Thrown when the text is null or empty</exception>
        public string Text 
        { 
            get { return text; }
            set 
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new FormatException("The text of the phonetic run cannot be null or empty");
                }
                text = XmlUtils.SanitizeXmlValue(value);
            } 
        }
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
            Text = text;
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
