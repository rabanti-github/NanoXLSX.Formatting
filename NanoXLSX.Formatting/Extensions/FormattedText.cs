/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Linq;
using NanoXLSX.Colors;
using NanoXLSX.Extensions;
using NanoXLSX.Interfaces;
using NanoXLSX.Styles;
using NanoXLSX.Utils;
using NanoXLSX.Utils.Xml;
using static NanoXLSX.Styles.Font;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX
{
    /// <summary>
    /// Represents a formatted text entry in Excel shared strings, supporting rich text with multiple runs and inline styles.
    /// </summary>
    public class FormattedText : IFormattableText
    {
        #region constants
        private const string SI_TAG = "si";
        private const string R_TAG = "r";
        private const string T_TAG = "t";
        private const string RPR_TAG = "rPr";
        private const string RPH_TAG = "rPh";
        private const string PHONETIC_PR_TAG = "phoneticPr";
        private const string PRESERVE_ATTRIBUTE_NAME = "space";
        private const string PRESERVE_ATTRIBUTE_PREFIX = "xml";
        private const string PRESERVE_ATTRIBUTE_VALUE = "preserve";
        #endregion

        #region staticFields

        /// <summary>
        /// Style to be used for line breaks in formatted text.
        /// </summary>
        public static readonly Style LineBreakStyle;

        static FormattedText()
        {
            LineBreakStyle = new Style();
            LineBreakStyle.CurrentCellXf.Alignment = CellXf.TextBreakValue.WrapText;
        }
        #endregion

        #region privateFields
        private readonly List<TextRun> runs = new List<TextRun>();
        private readonly List<PhoneticRun> phoeticRuns = new List<PhoneticRun>();
        #endregion

        #region properties
        /// <summary>
        /// Gets or sets whether the runs should be rendered with text wrapping, if there are line breaks present.
        /// </summary>
        /// \remark <remarks>The actual style component, responsible for rendering wrapped texts within a cell, is only added when saving the workbook</remarks>
        public bool WrapText { get; set; }

        /// <summary>
        /// The list of text runs in this formatted text.
        /// </summary>
        public IReadOnlyList<TextRun> Runs => runs.AsReadOnly();
        /// <summary>
        /// The list of phonetic runs (Ruby text) in this formatted text.
        /// </summary>
        public IReadOnlyList<PhoneticRun> PhoneticRuns => phoeticRuns.AsReadOnly();
        /// <summary>
        /// Phonetic properties for the formatted text (Phonetic run / Ruby text).
        /// </summary>
        public PhoneticProperties PhoneticProperties { get; set; }

        /// <summary>
        /// Gets the plain text content by concatenating all runs.
        /// </summary>
        public string PlainText => string.Concat(runs.Select(r => r.Text));
        #endregion

        #region methods

        /// <summary>
        /// Adds a text run with the specified style.
        /// </summary>
        /// <param name="text">The text content of the run</param>
        /// <param name="fontStyle">The font style to apply to the run (optional)</param>
        /// <returns>The current FormattedText instance for method chaining</returns>
        public FormattedText AddRun(string text, Font fontStyle = null)
        {
            if (string.IsNullOrEmpty(text))
            {
                throw new FormatException("Text cannot be null or empty");
            }
            CheckWrapText(text);
            runs.Add(new TextRun(text, fontStyle));
            return this;
        }

        /// <summary>
        /// Adds a text run using a style builder for inline configuration.
        /// </summary>
        /// <param name="text">The text content of the run</param>
        /// <param name="styleBuilder">An action to configure the inline style using an InlineStyleBuilder</param>
        /// <returns>The current FormattedText instance for method chaining</returns>
        public FormattedText AddRun(string text, Action<InlineStyleBuilder> styleBuilder)
        {
            if (string.IsNullOrEmpty(text))
            {
                throw new FormatException("Text cannot be null or empty");
            }
            CheckWrapText(text);
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
            WrapText = true;
        }

        /// <summary>
        /// Adds a phonetic run for pronunciation guidance (Ruby text, like Furigana, Pinyin or Zhuyin).
        /// </summary>
        /// <param name="text">The phonetic text (Ruby text)</param>
        /// <param name="startBase">The start index of the base text (character where the Ruby text starts) </param>
        /// <param name="endBase">The end index of the base text (character where the Ruby text ends)</param>
        /// <returns>The current FormattedText instance for method chaining</returns>
        public FormattedText AddPhoneticRun(string text, uint startBase, uint endBase)
        {
            phoeticRuns.Add(new PhoneticRun(text, startBase, endBase));
            return this;
        }

        /// <summary>
        /// Sets the phonetic properties for this formatted text, applied to the phonetic run (Ruby text).
        /// </summary>
        /// <param name="fontReference">The font reference that is used to render the phonetic run (Ruby text)</param>
        /// <param name="type">The phonetic type</param>
        /// <param name="alignment">The phonetic alignment</param>
        /// <returns>The current FormattedText instance for method chaining</returns>
        public FormattedText SetPhoneticProperties(Font fontReference, PhoneticRun.PhoneticType type = PhoneticRun.PhoneticType.FullwidthKatakana, PhoneticRun.PhoneticAlignment alignment = PhoneticRun.PhoneticAlignment.Left)
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

        /// <summary>
        /// Creates a deep copy of this FormattedText instance.
        /// </summary>
        /// <returns>The copied FormattedText instance</returns>
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
        /// Gets the string representation of the formatted text without formatting (plain text). The method is synonymous to the <see cref="PlainText"/> property.
        /// </summary>
        /// <returns>Plain text of the formatted text</returns>
        override public string ToString()
        {
            return PlainText;
        }

        #endregion

        #region privateMethods

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

            foreach (TextRun run in Runs)
            {
                XmlElement rElement = siElement.AddChildElement(R_TAG);

                if (run.FontStyle != null)
                {
                    XmlElement rPrElement = CreateRunPropertiesElement(run.FontStyle);
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
        /// Checks the text for line breaks and sets WrapText to true if any are found.
        /// </summary>
        /// <param name="text">Text to check</param>
        private void CheckWrapText(string text)
        {
            if (text.Contains("\n"))
            {
                WrapText = true;
            }
        }

        /// <summary>
        /// Get the XmlElement (interface implementation)
        /// </summary>
        /// <returns>XmlElement instance</returns>
        XmlElement IFormattableText.GetXmlElement()
        {
            return GetXmlElement();
        }



        /// <summary>
        /// Creates a phonetic properties element (&lt;phoneticPr&gt;)
        /// </summary>
        /// <returns>XmlElement instance</returns>
        private static XmlElement CreatePhoneticPropertiesElement(PhoneticProperties properties)
        {
            XmlElement phoneticPrElement = XmlElement.CreateElement(PHONETIC_PR_TAG);
            phoneticPrElement.AddAttribute("fontId", ParserUtils.ToString(properties.FontId));
            string typeValue = GetPhoneticTypeValue(properties.Type);
            phoneticPrElement.AddAttribute("type", typeValue);
            string alignmentValue = GetPhoneticAlignmentValue(properties.Alignment);
            phoneticPrElement.AddAttribute("alignment", alignmentValue);
            return phoneticPrElement;
        }


        /// <summary>
        /// Creates a run properties element (&lt;rPr&gt;) from an InlineStyle
        /// </summary>
        /// <param name="fontStyle">Font style instance</param>
        /// <returns>XmlElement instance</returns>
        private static XmlElement CreateRunPropertiesElement(Font fontStyle)
        {
            XmlElement rPrElement = XmlElement.CreateElement(RPR_TAG);
            if (!string.IsNullOrEmpty(fontStyle.Name))
            {
                rPrElement.AddChildElementWithAttribute("rFont", "val", fontStyle.Name);
            }
            if (fontStyle.Charset != Font.CharsetValue.Default)
            {
                rPrElement.AddChildElementWithAttribute("charset", "val", ParserUtils.ToString((int)fontStyle.Charset));
            }
            if (fontStyle.Family != Font.DefaultFontFamily)
            {
                rPrElement.AddChildElementWithAttribute("family", "val", ParserUtils.ToString((int)fontStyle.Family));
            }
            if (fontStyle.Bold)
            {
                rPrElement.AddChildElement("b");
            }
            if (fontStyle.Italic)
            {
                rPrElement.AddChildElement("i");
            }
            if (fontStyle.Strike)
            {
                rPrElement.AddChildElement("strike");
            }
            if (fontStyle.Outline)
            {
                rPrElement.AddChildElement("outline");
            }
            if (fontStyle.Shadow)
            {
                rPrElement.AddChildElement("shadow");
            }
            if (fontStyle.Condense)
            {
                rPrElement.AddChildElement("condense");
            }
            if (fontStyle.Extend)
            {
                rPrElement.AddChildElement("extend");
            }
            if (fontStyle.ColorValue != null && fontStyle.ColorValue.IsDefined)
            {
                XmlElement colorElement = CreateColorElement(fontStyle.ColorValue);
                rPrElement.AddChildElement(colorElement);
            }
            if (fontStyle.Size != Font.DefaultFontSize)
            {
                rPrElement.AddChildElementWithAttribute("sz", "val", ParserUtils.ToString(fontStyle.Size));
            }
            if (fontStyle.Underline != UnderlineValue.None && fontStyle.Underline != UnderlineValue.Single)
            {
                rPrElement.AddChildElementWithAttribute("u", "val", Font.GetUnderlineName(fontStyle.Underline));
            }
            else if (fontStyle.Underline == UnderlineValue.Single)
            {
                rPrElement.AddChildElement("u");
            }
            if (fontStyle.VerticalAlign != Font.VerticalTextAlignValue.None)
            {
                string vertAlignValue = Font.GetVerticalTextAlignName(fontStyle.VerticalAlign);
                rPrElement.AddChildElementWithAttribute("vertAlign", "val", vertAlignValue);
            }
            if (fontStyle.Scheme != Font.SchemeValue.None)
            {
                if (fontStyle.Scheme == SchemeValue.Major)
                {
                    rPrElement.AddChildElementWithAttribute("scheme", "val", "major");
                }
                else if (fontStyle.Scheme == SchemeValue.Minor)
                { rPrElement.AddChildElementWithAttribute("scheme", "val", "minor"); }
            }
            return rPrElement;
        }

        /// <summary>
        /// Creates a text element (&lt;t&gt;) with proper whitespace preservation
        /// </summary>
        /// <param name="text">Text content</param>
        /// <returns>XmlElement instance</returns>
        /// \remark <remarks>Null or empty values must be removed in advance</remarks>
        private static XmlElement CreateTextElement(string text)
        {
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
        /// Creates a color element (&lt;color&gt;) from a Color instance
        /// </summary>
        /// <param name="color">Color instance</param>
        /// <returns>XmlElement instance</returns>
        private static XmlElement CreateColorElement(Color color)
        {
            XmlElement colorElement = XmlElement.CreateElement("color");
            if (color.Value is AutoColor)
            {
                colorElement.AddAttribute("auto", "1");
            }
            else if (color.Value is IndexedColor)
            {
                colorElement.AddAttribute("indexed", color.Value.StringValue);
            }
            else if (color.Value is SrgbColor)
            {
                colorElement.AddAttribute("rgb", color.Value.StringValue);
            }
            else if (color.Value is ThemeColor)
            {
                colorElement.AddAttribute("theme", color.Value.StringValue);
            }
            else if (color.Value is SystemColor)
            {
                colorElement.AddAttribute("system", color.Value.StringValue);
            }
            if (color.Tint.HasValue)
            {
                colorElement.AddAttribute("tint", ParserUtils.ToString(color.Tint.Value));
            }
            return colorElement;
        }

        /// <summary>
        /// Creates a phonetic run element (&lt;rPh&gt;)
        /// </summary>
        /// <param name="phoneticRun" >PhoneticRun instance</param>
        /// <returns>XmlElement instance</returns>
        private static XmlElement CreatePhoneticRunElement(PhoneticRun phoneticRun)
        {
            XmlElement rPhElement = XmlElement.CreateElement(RPH_TAG);
            rPhElement.AddAttribute("sb", ParserUtils.ToString(phoneticRun.StartBase));
            rPhElement.AddAttribute("eb", ParserUtils.ToString(phoneticRun.EndBase));
            XmlElement tElement = CreateTextElement(phoneticRun.Text);
            rPhElement.AddChildElement(tElement);
            return rPhElement;
        }

        /// <summary>
        /// Converts PhoneticType enum to OOXML string value
        /// </summary>
        /// <param name="type">PhoneticType enum value</param>
        /// <returns>OOXML string value</returns>
        private static string GetPhoneticTypeValue(PhoneticRun.PhoneticType type)
        {
            switch (type)
            {
                case PhoneticRun.PhoneticType.HalfwidthKatakana:
                    return "halfwidthKatakana";
                case PhoneticRun.PhoneticType.Hiragana:
                    return "Hiragana";
                case PhoneticRun.PhoneticType.NoConversion:
                    return "noConversion";
                default:
                    return "fullwidthKatakana";
            }
        }

        /// <summary>
        /// Converts PhoneticAlignment enum to OOXML string value
        /// </summary>
        /// <param name="alignment">PhoneticAlignment enum value</param>
        /// <returns>OOXML string value</returns>   
        private static string GetPhoneticAlignmentValue(PhoneticRun.PhoneticAlignment alignment)
        {
            switch (alignment)
            {
                case PhoneticRun.PhoneticAlignment.NoControl:
                    return "noControl";
                case PhoneticRun.PhoneticAlignment.Center:
                    return "center";
                case PhoneticRun.PhoneticAlignment.Distributed:
                    return "distributed";
                default:
                    return "left";
            }
        }

        /// <summary>
        /// Equals method override for comparing two FormattedText instances.
        /// </summary>
        /// <param name="obj">Other object to compare</param>
        /// <returns>True, if the other object is equal</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is FormattedText text))
                return false;

            return runs.SequenceEqual(text.runs) &&
                   phoeticRuns.SequenceEqual(text.phoeticRuns) &&
                   WrapText == text.WrapText &&
                   EqualityComparer<PhoneticProperties>.Default.Equals(PhoneticProperties, text.PhoneticProperties);
        }

        /// <summary>
        /// GetHashCode method override for FormattedText.
        /// </summary>
        /// <returns>Hash code of the current object</returns>
        public override int GetHashCode()
        {
            var hashCode = 703246462;
            foreach (var run in runs)
            {
                hashCode = hashCode * -1521134295 + (run?.GetHashCode() ?? 0);
            }
            foreach (var phoneticRun in phoeticRuns)
            {
                hashCode = hashCode * -1521134295 + (phoneticRun?.GetHashCode() ?? 0);
            }
            hashCode = hashCode * -1521134295 + WrapText.GetHashCode();
            hashCode = hashCode * -1521134295 + (PhoneticProperties?.GetHashCode() ?? 0);
            return hashCode;
        }


        #endregion
    }
}
