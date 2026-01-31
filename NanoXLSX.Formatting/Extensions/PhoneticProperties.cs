/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
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
        /// Internal Id of the font used for phonetic text
        /// </summary>
        /// \remark <remarks>This ID will be resolved automatically, when the workbook is saved, according to <see cref="PhoneticProperties.FontReference"/></remarks>
        public int FontId { get; set; }
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

        /// <summary>
        /// Equals override to compare phonetic properties.
        /// </summary>
        /// <param name="obj">Other object to compare</param>
        /// <returns>True, if the other object is equal</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is PhoneticProperties properties)) { return false; }
            return ((FontReference == null && properties.FontReference == null) ||
                    (FontReference != null && FontReference.Equals(properties.FontReference))) &&
                   Type == properties.Type &&
                   Alignment == properties.Alignment;
        }

        /// <summary>
        /// HashCode override for phonetic properties.
        /// </summary>
        /// <returns>Hash code of the current object</returns>
        public override int GetHashCode()
        {
            var hashCode = 136542229;
            hashCode = hashCode * -1521134295 + EqualityComparer<Font>.Default.GetHashCode(FontReference);
            hashCode = hashCode * -1521134295 + Type.GetHashCode();
            hashCode = hashCode * -1521134295 + Alignment.GetHashCode();
            return hashCode;
        }
    }
}
