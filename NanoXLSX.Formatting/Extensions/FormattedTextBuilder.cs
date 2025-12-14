using System;
using System.Collections.Generic;
using System.Text;

namespace NanoXLSX.Extensions
{
    /// <summary>
    /// Builder for creating formatted text entries with a fluent API.
    /// </summary>
    public class FormattedTextBuilder
    {
        private readonly FormattedText _formattedText = new FormattedText();

        public FormattedTextBuilder AddRun(string text, InlineStyle style = null)
        {
            _formattedText.AddRun(text, style);
            return this;
        }

        public FormattedTextBuilder AddRun(string text, Action<InlineStyleBuilder> styleBuilder)
        {
            _formattedText.AddRun(text, styleBuilder);
            return this;
        }

        public FormattedTextBuilder AddPhoneticRun(string text, uint startBase, uint endBase)
        {
            _formattedText.AddPhoneticRun(text, startBase, endBase);
            return this;
        }

        public FormattedTextBuilder SetPhoneticProperties(object fontReference, PhoneticType type = PhoneticType.FullwidthKatakana, PhoneticAlignment alignment = PhoneticAlignment.Left)
        {
            _formattedText.SetPhoneticProperties(fontReference, type, alignment);
            return this;
        }

        public FormattedText Build()
        {
            return _formattedText;
        }

        public static implicit operator FormattedText(FormattedTextBuilder builder)
        {
            return builder.Build();
        }
    }
}
