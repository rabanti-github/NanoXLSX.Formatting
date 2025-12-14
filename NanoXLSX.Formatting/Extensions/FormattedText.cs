using System;
using System.Collections.Generic;
using System.Linq;
using NanoXLSX.Extensions;
using NanoXLSX.Interfaces;
using NanoXLSX.Utils;
using NanoXLSX.Utils.Xml;

namespace NanoXLSX
{
    /// <summary>
    /// Represents a formatted text entry in Excel shared strings, supporting rich text with multiple runs and inline styles.
    /// </summary>
    public class FormattedText : IFormattableText
    {
        private const string SI_TAG = "si";
        private const string R_TAG = "r";
        private const string T_TAG = "t";
        private const string RPR_TAG = "rPr";
        private const string RPH_TAG = "rPh";
        private const string PHONETIC_PR_TAG = "phoneticPr";
        private const string PRESERVE_ATTRIBUTE_NAME = "space";
        private const string PRESERVE_ATTRIBUTE_PREFIX = "xml";
        private const string PRESERVE_ATTRIBUTE_VALUE = "preserve";

        public int FontId { get; set; }


        private readonly List<TextRun> runs = new List<TextRun>();
        private readonly List<PhoneticRun> phoeticRuns = new List<PhoneticRun>();

        public IReadOnlyList<TextRun> Runs => runs.AsReadOnly();
        public IReadOnlyList<PhoneticRun> PhoneticRuns => phoeticRuns.AsReadOnly();
        public PhoneticProperties PhoneticProperties { get; set; }

        /// <summary>
        /// Gets the plain text content by concatenating all runs.
        /// </summary>
        public string PlainText => string.Concat(runs.Select(r => r.Text));

        /// <summary>
        /// Adds a text run with the specified style.
        /// </summary>
        public FormattedText AddRun(string text, InlineStyle style = null)
        {
            if (string.IsNullOrEmpty(text))
            {
                throw new ArgumentException("Text cannot be null or empty.", nameof(text));
            }

            runs.Add(new TextRun(text, style));
            return this;
        }

        /// <summary>
        /// Adds a text run using a style builder for inline configuration.
        /// </summary>
        public FormattedText AddRun(string text, Action<InlineStyleBuilder> styleBuilder)
        {
            if (string.IsNullOrEmpty(text))
                throw new ArgumentException("Text cannot be null or empty.", nameof(text));

            var builder = new InlineStyleBuilder();
            styleBuilder?.Invoke(builder);
            runs.Add(new TextRun(text, builder.Build()));
            return this;
        }

        /// <summary>
        /// Adds a line break to the formatted text. By default, the last run's style is used.
        /// </summary>
        /// <param name="useStyleFromLastRun">If set to false, a new run without style will be created for the line break</param>
        public void AddLineBreak(bool useStyleFromLastRun = true)
        {
            if (runs.Count == 0 || !useStyleFromLastRun)
            {
                runs.Add(new TextRun(Environment.NewLine, null));
            }
            else
            {
                runs[runs.Count - 1].Text += Environment.NewLine;
            }
        }

        /// <summary>
        /// Adds a phonetic run for pronunciation guidance.
        /// </summary>
        public FormattedText AddPhoneticRun(string text, uint startBase, uint endBase)
        {
            phoeticRuns.Add(new PhoneticRun(text, startBase, endBase));
            return this;
        }

        /// <summary>
        /// Sets the phonetic properties for this formatted text.
        /// </summary>
        public FormattedText SetPhoneticProperties(object fontReference, PhoneticType type = PhoneticType.FullwidthKatakana, PhoneticAlignment alignment = PhoneticAlignment.Left)
        {
            PhoneticProperties = new PhoneticProperties(fontReference)
            {
                Type = type,
                Alignment = alignment
            };
            return this;
        }

        /// <summary>
        /// Clears all runs from this formatted text.
        /// </summary>
        public void Clear()
        {
            runs.Clear();
            phoeticRuns.Clear();
            PhoneticProperties = null;
        }

        public FormattedText Copy()
        {
            FormattedText copy = new FormattedText();
            foreach (TextRun run in runs)
            {
                copy.runs.Add(run.Copy());
            }
            foreach (PhoneticRun phoneticRun in phoeticRuns)
            {
                copy.phoeticRuns.Add(phoneticRun.Copy());
            }
            if (PhoneticProperties != null)
            {
                copy.PhoneticProperties = PhoneticProperties.Copy();
            }
            return copy;
        }

        /// <summary>
        /// Get the XmlElement (interface implementation)
        /// </summary>
        /// <returns>XmlElement instance</returns>
        internal XmlElement GetXmlElement()
        {
            XmlElement siElement = XmlElement.CreateElement(SI_TAG);

            if (Runs.Count == 0)
            {
                return siElement;
            }

            foreach (Extensions.TextRun run in Runs)
            {
                XmlElement rElement = siElement.AddChildElement(R_TAG);

                if (run.Style != null)
                {
                    XmlElement rPrElement = CreateRunPropertiesElement(run.Style);
                    rElement.AddChildElement(rPrElement);
                }

                XmlElement tElement = CreateTextElement(run.Text);
                rElement.AddChildElement(tElement);
            }

            foreach (var phoneticRun in PhoneticRuns)
            {
                XmlElement rPhElement = CreatePhoneticRunElement(phoneticRun);
                siElement.AddChildElement(rPhElement);
            }

            if (PhoneticProperties != null)
            {
                XmlElement phoneticPrElement = CreatePhoneticPropertiesElement(PhoneticProperties);
                siElement.AddChildElement(phoneticPrElement);
            }

            return siElement;
        }

        /// <summary>
        /// Creates a text element (&lt;t&gt;) with proper whitespace preservation
        /// </summary>
        private XmlElement CreateTextElement(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return XmlElement.CreateElement(T_TAG);
            }

            string value = XmlUtils.SanitizeXmlValue(text);
            value = ParserUtils.NormalizeNewLines(value);

            XmlElement element;
            if (char.IsWhiteSpace(value, 0) || char.IsWhiteSpace(value, value.Length - 1))
            {
                element = XmlElement.CreateElementWithAttribute(
                    T_TAG,
                    PRESERVE_ATTRIBUTE_NAME,
                    PRESERVE_ATTRIBUTE_VALUE,
                    "",
                    PRESERVE_ATTRIBUTE_PREFIX);
            }
            else
            {
                element = XmlElement.CreateElement(T_TAG);
            }

            element.InnerValue = value;
            return element;
        }

        /// <summary>
        /// Creates a run properties element (&lt;rPr&gt;) from an InlineStyle
        /// </summary>
        private XmlElement CreateRunPropertiesElement(InlineStyle style)
        {
            XmlElement rPrElement = XmlElement.CreateElement(RPR_TAG);

            if (!string.IsNullOrEmpty(style.FontName))
            {
                rPrElement.AddChildElementWithAttribute("rFont", "val", style.FontName);
            }

            if (style.Charset.HasValue)
            {
                rPrElement.AddChildElementWithAttribute("charset", "val", ParserUtils.ToString((int)style.Charset.Value));
            }

            if (style.Family.HasValue)
            {
                rPrElement.AddChildElementWithAttribute("family", "val", ParserUtils.ToString((int)style.Family.Value));
            }

            if (style.Bold.HasValue && style.Bold.Value)
            {
                rPrElement.AddChildElement("b");
            }

            if (style.Italic.HasValue && style.Italic.Value)
            {
                rPrElement.AddChildElement("i");
            }

            if (style.Strikethrough.HasValue && style.Strikethrough.Value)
            {
                rPrElement.AddChildElement("strike");
            }

            if (style.Outline.HasValue && style.Outline.Value)
            {
                rPrElement.AddChildElement("outline");
            }

            if (style.Shadow.HasValue && style.Shadow.Value)
            {
                rPrElement.AddChildElement("shadow");
            }

            if (style.Condense.HasValue && style.Condense.Value)
            {
                rPrElement.AddChildElement("condense");
            }

            if (style.Extend.HasValue && style.Extend.Value)
            {
                rPrElement.AddChildElement("extend");
            }

            if (style.Color != null)
            {
                XmlElement colorElement = CreateColorElement(style.Color);
                rPrElement.AddChildElement(colorElement);
            }

            if (style.FontSize.HasValue)
            {
                rPrElement.AddChildElementWithAttribute("sz", "val", ParserUtils.ToString(style.FontSize.Value));
            }

            if (style.Underline.HasValue)
            {
                string underlineValue = GetUnderlineValue(style.Underline.Value);
                if (underlineValue != "single")
                {
                    rPrElement.AddChildElementWithAttribute("u", "val", underlineValue);
                }
                else
                {
                    rPrElement.AddChildElement("u");
                }
            }

            if (style.VerticalAlign.HasValue)
            {
                string vertAlignValue = GetVerticalAlignmentValue(style.VerticalAlign.Value);
                rPrElement.AddChildElementWithAttribute("vertAlign", "val", vertAlignValue);
            }

            if (style.Scheme.HasValue)
            {
                string schemeValue = GetFontSchemeValue(style.Scheme.Value);
                rPrElement.AddChildElementWithAttribute("scheme", "val", schemeValue);
            }

            return rPrElement;
        }

        /// <summary>
        /// Creates a color element (&lt;color&gt;) from a Color instance
        /// </summary>
        private XmlElement CreateColorElement(Extensions.Color color)
        {
            XmlElement colorElement = XmlElement.CreateElement("color");

            if (color.Auto.HasValue && color.Auto.Value)
            {
                colorElement.AddAttribute("auto", "1");
            }

            if (color.Indexed.HasValue)
            {
                colorElement.AddAttribute("indexed", ParserUtils.ToString(color.Indexed.Value));
            }

            if (!string.IsNullOrEmpty(color.Rgb))
            {
                colorElement.AddAttribute("rgb", color.Rgb);
            }

            if (color.Theme.HasValue)
            {
                colorElement.AddAttribute("theme", color.Theme.Value.ToString());
            }

            if (color.Tint != 0.0)
            {
                colorElement.AddAttribute("tint", color.Tint.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }

            return colorElement;
        }

        /// <summary>
        /// Creates a phonetic run element (&lt;rPh&gt;)
        /// </summary>
        private XmlElement CreatePhoneticRunElement(PhoneticRun phoneticRun)
        {
            XmlElement rPhElement = XmlElement.CreateElement(RPH_TAG);
            rPhElement.AddAttribute("sb", phoneticRun.StartBase.ToString());
            rPhElement.AddAttribute("eb", phoneticRun.EndBase.ToString());

            XmlElement tElement = CreateTextElement(phoneticRun.Text);
            rPhElement.AddChildElement(tElement);

            return rPhElement;
        }

        /// <summary>
        /// Creates a phonetic properties element (&lt;phoneticPr&gt;)
        /// </summary>
        private XmlElement CreatePhoneticPropertiesElement(PhoneticProperties properties)
        {
            XmlElement phoneticPrElement = XmlElement.CreateElement(PHONETIC_PR_TAG);
            phoneticPrElement.AddAttribute("fontId", ParserUtils.ToString(FontId));

            if (properties.Type != PhoneticType.FullwidthKatakana)
            {
                string typeValue = GetPhoneticTypeValue(properties.Type);
                phoneticPrElement.AddAttribute("type", typeValue);
            }

            if (properties.Alignment != PhoneticAlignment.Left)
            {
                string alignmentValue = GetPhoneticAlignmentValue(properties.Alignment);
                phoneticPrElement.AddAttribute("alignment", alignmentValue);
            }

            return phoneticPrElement;
        }

        /// <summary>
        /// Converts UnderlineStyle enum to OOXML string value
        /// </summary>
        private string GetUnderlineValue(UnderlineStyle style)
        {
            switch (style)
            {
                case UnderlineStyle.Single:
                    return "single";
                case UnderlineStyle.Double:
                    return "double";
                case UnderlineStyle.SingleAccounting:
                    return "singleAccounting";
                case UnderlineStyle.DoubleAccounting:
                    return "doubleAccounting";
                case UnderlineStyle.None:
                    return "none";
                default:
                    return "single";
            }
        }

        /// <summary>
        /// Converts VerticalAlignment enum to OOXML string value
        /// </summary>
        private string GetVerticalAlignmentValue(VerticalAlignment alignment)
        {
            switch (alignment)
            {
                case VerticalAlignment.Baseline:
                    return "baseline";
                case VerticalAlignment.Superscript:
                    return "superscript";
                case VerticalAlignment.Subscript:
                    return "subscript";
                default:
                    return "baseline";
            }
        }

        /// <summary>
        /// Converts FontScheme enum to OOXML string value
        /// </summary>
        private string GetFontSchemeValue(FontScheme scheme)
        {
            switch (scheme)
            {
                case FontScheme.None:
                    return "none";
                case FontScheme.Major:
                    return "major";
                case FontScheme.Minor:
                    return "minor";
                default:
                    return "none";
            }
        }

        /// <summary>
        /// Converts PhoneticType enum to OOXML string value
        /// </summary>
        private string GetPhoneticTypeValue(PhoneticType type)
        {
            switch (type)
            {
                case PhoneticType.HalfwidthKatakana:
                    return "halfwidthKatakana";
                case PhoneticType.FullwidthKatakana:
                    return "fullwidthKatakana";
                case PhoneticType.Hiragana:
                    return "Hiragana";
                case PhoneticType.NoConversion:
                    return "noConversion";
                default:
                    return "fullwidthKatakana";
            }
        }

        /// <summary>
        /// Converts PhoneticAlignment enum to OOXML string value
        /// </summary>
        private string GetPhoneticAlignmentValue(PhoneticAlignment alignment)
        {
            switch (alignment)
            {
                case PhoneticAlignment.NoControl:
                    return "noControl";
                case PhoneticAlignment.Left:
                    return "left";
                case PhoneticAlignment.Center:
                    return "center";
                case PhoneticAlignment.Distributed:
                    return "distributed";
                default:
                    return "left";
            }
        }

        XmlElement IFormattableText.GetXmlElement()
        {
            return GetXmlElement();
        }
    }
}
