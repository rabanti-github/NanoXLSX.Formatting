/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using NanoXLSX.Styles;

namespace NanoXLSX.Extensions
{
    /// <summary>
    /// Represents phonetic properties, applied to the phonetic runs of a formatted text.
    /// </summary>
    public class PhoneticProperties
    {
        /// <summary>
        /// Font reference that is used to render the Ruby text
        /// </summary>
        public Font FontReference { get; set; }
        /// <summary>
        /// The type of phonetic text representation. Default is FullwidthKatakana.
        /// </summary>
        public PhoneticRun.PhoneticType Type { get; set; } = PhoneticRun.PhoneticType.FullwidthKatakana;
        /// <summary>
        /// Alignment of the phonetic text relative to the base text. Default is Left.
        /// </summary>
        public PhoneticRun.PhoneticAlignment Alignment { get; set; } = PhoneticRun.PhoneticAlignment.Left;

        /// <summary>
        /// Default constructor to create phonetic properties, using the specified font reference.
        /// </summary>
        /// <param name="fontReference">Font reference. If null, a default font reference number will be selected, probably 0</param>
        public PhoneticProperties(Font fontReference)
        {
            FontReference = fontReference; // Validation will be performed when XLSX file is generated
        }

        /// <summary>
        /// Creates a copy of the current phonetic properties.
        /// </summary>
        /// <returns>Copy of the phonetic properties</returns>
        public PhoneticProperties Copy()
        {
            return new PhoneticProperties(FontReference)
            {
                Type = this.Type,
                Alignment = this.Alignment
            };
        }

    }
}
