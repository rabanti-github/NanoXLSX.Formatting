using System;
using System.Collections.Generic;
using System.Text;

namespace NanoXLSX.Extensions
{
    /// <summary>
    /// Represents a phonetic run that provides pronunciation guidance for text.
    /// </summary>
    public class PhoneticRun
    {
        public string Text { get; set; }
        public uint StartBase { get; set; }
        public uint EndBase { get; set; }

        public PhoneticRun(string text, uint startBase, uint endBase)
        {
            Text = text ?? throw new ArgumentNullException(nameof(text));
            StartBase = startBase;
            EndBase = endBase;
        }

        public PhoneticRun Copy()
        {
            return new PhoneticRun(this.Text, this.StartBase, this.EndBase);
        }
    }
}
