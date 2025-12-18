using System;
using System.Collections.Generic;
using System.Text;
using NanoXLSX.Styles;

namespace NanoXLSX.Extensions
{
    /// <summary>
    /// Represents phonetic properties for the entire formatted text.
    /// </summary>
    public class PhoneticProperties
    {
        public Font FontReference { get; set; }
        public PhoneticType Type { get; set; } = PhoneticType.FullwidthKatakana;
        public PhoneticAlignment Alignment { get; set; } = PhoneticAlignment.Left;

        public PhoneticProperties(Font fontReference)
        {
            FontReference = fontReference ?? throw new ArgumentNullException(nameof(fontReference));
        }

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
