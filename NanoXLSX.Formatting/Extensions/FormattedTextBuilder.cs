/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using NanoXLSX.Styles;

namespace NanoXLSX.Extensions
{
    /// <summary>
    /// Builder for creating formatted text entries with a fluent API.
    /// </summary>
    public class FormattedTextBuilder
    {
        private readonly FormattedText formattedText = new FormattedText();

        /// <summary>
        /// Adds a text run with an optional style to the formatted text.
        /// </summary>
        /// <param name="text">Plain text of the run</param>
        /// <param name="font">Font style</param>
        /// <returns>The current builder instance</returns>
        public FormattedTextBuilder AddRun(string text, Font font = null)
        {
            formattedText.AddRun(text, font);
            return this;
        }

        /// <summary>
        /// Adds a text run with a style defined by a style builder action.
        /// </summary>
        /// <param name="text">Plain text of the run</param>
        /// <param name="styleBuilder">Action to build the inline style</param>
        /// <returns>The current builder instance</returns>
        public FormattedTextBuilder AddRun(string text, Action<InlineStyleBuilder> styleBuilder)
        {
            formattedText.AddRun(text, styleBuilder);
            return this;
        }

        /// <summary>
        /// Adds a phonetic run to the formatted text.
        /// </summary>
        /// <param name="text">The phonetic text to be displayed (Ruby text,like Furigana, Pinyin or Zhuyin)</param>
        /// <param name="startBase">The start index of the base text (character where the Ruby text starts)</param>
        /// <param name="endBase">The end index of the base text (character where the Ruby text ends)</param>
        /// <returns>The current builder instance</returns>
        public FormattedTextBuilder AddPhoneticRun(string text, uint startBase, uint endBase)
        {
            formattedText.AddPhoneticRun(text, startBase, endBase);
            return this;
        }

        /// <summary>
        /// Sets the phonetic properties for the formatted text.
        /// </summary>
        /// <param name="fontReference">Font reference that is used to render the Ruby text</param>
        /// <param name="type">Phonetic type</param>
        /// <param name="alignment">Phonetic alignment</param>
        /// <returns>The current builder instance</returns>
        public FormattedTextBuilder SetPhoneticProperties(Font fontReference, PhoneticRun.PhoneticType type = PhoneticRun.PhoneticType.FullwidthKatakana, PhoneticRun.PhoneticAlignment alignment = PhoneticRun.PhoneticAlignment.Left)
        {
            formattedText.SetPhoneticProperties(fontReference, type, alignment);
            return this;
        }

        /// <summary>
        /// Method to build the formatted text instance.
        /// </summary>
        /// <returns>The constructed FormattedText instance</returns>
        public FormattedText Build()
        {
            return formattedText;
        }

        /// <summary>
        /// Implicit conversion operator to convert the builder to a FormattedText instance.
        /// </summary>
        /// <param name="builder">The FormattedTextBuilder instance</param>
        public static implicit operator FormattedText(FormattedTextBuilder builder)
        {
            return builder.Build();
        }
    }
}
