/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using NanoXLSX.Colors;
using NanoXLSX.Extensions;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Styles;
using NanoXLSX.Utils;
using NanoXLSX.Utils.Xml;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class implementing a reader for the shared strings table of XLSX files, with support for formatted text and phonetic characters.
    /// </summary>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.SharedStringsReader, PlugInOrder = 1000)]
    internal class FormattedSharedStringsReader : ISharedStringReader
    {
        /// <summary>
        /// Entity id to identify auxiliary data
        /// </summary>
        internal static readonly int AUXILIARY_DATA_ID = 854563987;

        #region privateFields
        private bool capturePhoneticCharacters;
        private readonly List<PhoneticInfo> phoneticsInfo;
        private Stream stream;
        private Workbook workbook;
        #endregion

        #region properties

        /// <summary>
        /// List of shared string entries
        /// </summary>
        /// <value>
        /// String entry, sorted by its internal index of the table
        /// </value>
        public List<string> SharedStrings { get; private set; }

        /// <summary>
        /// Dictionary of formatted text entries, indexed by their plain text value (to be replaced in the workbook)
        /// </summary>
        public Dictionary<string, FormattedText> FormattedTexts { get; private set; }

        /// <summary>
        /// Workbook reference where read data is stored (should not be null)
        /// </summary>
        public Workbook Workbook { get => workbook; set => workbook = value; }
        /// <summary>
        /// Reader options
        /// </summary>
        public IOptions Options { get; set; }
        /// <summary>
        /// Reference to a reader plug-in handler, to be used for post operations in the <see cref="Execute"/> method
        /// </summary>
        public Action<Stream, Workbook, string, IOptions, int?> InlinePluginHandler { get; set; }
        #endregion

        #region constructors
        /// <summary>
        /// Default constructor - Must be defined for instantiation of the plug-ins
        /// </summary>
        public FormattedSharedStringsReader()
        {
            phoneticsInfo = new List<PhoneticInfo>();
            SharedStrings = new List<string>();
            FormattedTexts = new Dictionary<string, FormattedText>();
        }
        #endregion

        #region methods
        /// <summary>
        /// Initialization method (interface implementation)
        /// </summary>
        /// <param name="stream">Stream to be read</param>
        /// <param name="workbook">Workbook reference</param>
        /// <param name="readerOptions">Reader options</param>
        /// <param name="inlinePluginHandler">Inline plug-in handler</param>
        public void Init(Stream stream, Workbook workbook, IOptions readerOptions, Action<Stream, Workbook, string, IOptions, int?> inlinePluginHandler)
        {
            this.stream = stream;
            this.workbook = workbook;
            if (readerOptions is ITextOptions options)
            {
                this.capturePhoneticCharacters = options.EnforcePhoneticCharacterImport;
            }
            this.InlinePluginHandler = inlinePluginHandler;
        }

        /// <summary>
        /// Method to execute the main logic of the plug-in (interface implementation)
        /// </summary>
        /// <exception cref="Exceptions.IOException">Throws an IOException in case of a error during reading</exception>
        public void Execute()
        {
            try
            {
                using (stream) // Close after processing
                {
                    bool hasFormattedText = false;
                    StringBuilder sb = new StringBuilder();
                    using (XmlReader reader = XmlReader.Create(stream, XmlStreamUtils.CreateSettings()))
                    {
                        while (reader.Read())
                        {
                            if (!XmlStreamUtils.IsElement(reader, "si"))
                                continue;

                            sb.Clear();
                            phoneticsInfo.Clear();

                            FormattedText formattedText;
                            using (XmlReader siReader = reader.ReadSubtree())
                            {
                                siReader.Read(); // consume the <si> open tag
                                formattedText = ProcessSharedStringItem(siReader, ref sb);
                            }

                            string textValue;
                            if (capturePhoneticCharacters)
                            {
                                textValue = ProcessPhoneticCharacters(sb);
                                formattedText.OverridePlainText(textValue);
                            }
                            else if (formattedText != null && string.IsNullOrEmpty(formattedText.PlainText) && sb.Length > 0)
                            {
                                textValue = sb.ToString();
                                formattedText.OverridePlainText(textValue); // Fallback to prevent data loss
                            }
                            else
                            {
                                textValue = sb.ToString();
                            }

                            if (formattedText != null)
                            {
                                string key = PlugInUUID.SharedStringsReader + textValue;
                                SharedStrings.Add(key);
                                FormattedTexts[key] = formattedText;
                                hasFormattedText = true;
                            }
                            else
                            {
                                SharedStrings.Add(textValue);
                            }
                        }
                    }
                    InlinePluginHandler?.Invoke(stream, Workbook, PlugInUUID.SharedStringsInlineReader, Options, null);
                    if (hasFormattedText)
                    {
                        Workbook.AuxiliaryData.SetData(PlugInUUID.SharedStringsReader, AUXILIARY_DATA_ID, FormattedTexts);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new IOException("The XML entry could not be read from the " + nameof(stream) + ". Please see the inner exception:", ex);
            }
        }

        /// <summary>
        /// Processes a shared string item (&lt;si&gt;) node and creates FormattedText if it contains runs or phonetic data
        /// </summary>
        /// <param name="siReader">XmlReader subtree positioned at the &lt;si&gt; open tag</param>
        /// <param name="sb">StringBuilder for plain text extraction</param>
        /// <returns>FormattedText if the item has formatting or phonetic data, null for plain text</returns>
        private FormattedText ProcessSharedStringItem(XmlReader siReader, ref StringBuilder sb)
        {
            bool hasRuns = false;
            bool hasFormattedContent = false;
            FormattedText formattedText = null;

            while (siReader.Read())
            {
                if (siReader.NodeType != XmlNodeType.Element)
                    continue;

                if (XmlStreamUtils.IsElement(siReader, "r"))
                {
                    hasRuns = true;
                    hasFormattedContent = true;
                    if (formattedText == null)
                        formattedText = new FormattedText();
                    using (XmlReader runReader = siReader.ReadSubtree())
                    {
                        runReader.Read(); // consume the <r> open tag
                        ProcessTextRun(runReader, formattedText, ref sb);
                    }
                }
                else if (XmlStreamUtils.IsElement(siReader, "rPh"))
                {
                    hasFormattedContent = true;
                    if (formattedText == null)
                        formattedText = new FormattedText();
                    using (XmlReader rPhReader = siReader.ReadSubtree())
                    {
                        rPhReader.Read(); // consume the <rPh> open tag
                        ProcessPhoneticRun(rPhReader, formattedText);
                    }
                }
                else if (XmlStreamUtils.IsElement(siReader, "phoneticPr"))
                {
                    hasFormattedContent = true;
                    if (formattedText == null)
                        formattedText = new FormattedText();
                    ProcessPhoneticProperties(siReader, formattedText);
                    using (siReader.ReadSubtree()) { } // consume element; positions reader at end element
                }
                else if (XmlStreamUtils.IsElement(siReader, "t") && !hasRuns)
                {
                    // Plain text or text accompanying only phonetic runs (no rich-text runs)
                    string text;
                    using (XmlReader tReader = siReader.ReadSubtree())
                    {
                        tReader.Read(); // position at <t>
                        text = tReader.ReadElementContentAsString();
                    }
                    sb.Append(text);
                }
            }

            return hasFormattedContent ? formattedText : null;
        }

        /// <summary>
        /// Processes a text run (&lt;r&gt;) subtree
        /// </summary>
        /// <param name="runReader">XmlReader subtree positioned at the &lt;r&gt; open tag</param>
        /// <param name="formattedText">FormattedText to add the run to</param>
        /// <param name="sb">StringBuilder for plain text extraction</param>
        private void ProcessTextRun(XmlReader runReader, FormattedText formattedText, ref StringBuilder sb)
        {
            Font fontStyle = null;
            string text = null;

            while (runReader.Read())
            {
                if (runReader.NodeType != XmlNodeType.Element)
                    continue;

                if (XmlStreamUtils.IsElement(runReader, "rPr"))
                {
                    using (XmlReader rPrReader = runReader.ReadSubtree())
                    {
                        rPrReader.Read(); // consume the <rPr> open tag
                        fontStyle = ParseRunProperties(rPrReader);
                    }
                }
                else if (XmlStreamUtils.IsElement(runReader, "t"))
                {
                    using (XmlReader tReader = runReader.ReadSubtree())
                    {
                        tReader.Read(); // position at <t>
                        text = tReader.ReadElementContentAsString();
                    }
                    sb.Append(text);
                }
            }

            if (!string.IsNullOrEmpty(text))
            {
                formattedText.AddRun(text, fontStyle);
            }
        }

        /// <summary>
        /// Parses run properties (&lt;rPr&gt;) subtree and creates a Font object
        /// </summary>
        /// <param name="rPrReader">XmlReader subtree positioned at the &lt;rPr&gt; open tag</param>
        /// <returns>Font object with parsed properties</returns>
        private Font ParseRunProperties(XmlReader rPrReader)
        {
            Font font = new Font();

            while (rPrReader.Read())
            {
                if (rPrReader.NodeType != XmlNodeType.Element)
                    continue;

                string nodeName = rPrReader.LocalName;

                if (nodeName.Equals("rFont", StringComparison.OrdinalIgnoreCase))
                {
                    font.Name = rPrReader.GetAttribute("val");
                }
                else if (nodeName.Equals("charset", StringComparison.OrdinalIgnoreCase))
                {
                    string val = rPrReader.GetAttribute("val");
                    if (!string.IsNullOrEmpty(val))
                        font.Charset = (Font.CharsetValue)ParserUtils.ParseInt(val);
                }
                else if (nodeName.Equals("family", StringComparison.OrdinalIgnoreCase))
                {
                    string val = rPrReader.GetAttribute("val");
                    if (!string.IsNullOrEmpty(val))
                        font.Family = (Font.FontFamilyValue)ParserUtils.ParseInt(val);
                }
                else if (nodeName.Equals("b", StringComparison.OrdinalIgnoreCase))
                {
                    font.Bold = true;
                }
                else if (nodeName.Equals("i", StringComparison.OrdinalIgnoreCase))
                {
                    font.Italic = true;
                }
                else if (nodeName.Equals("strike", StringComparison.OrdinalIgnoreCase))
                {
                    font.Strike = true;
                }
                else if (nodeName.Equals("outline", StringComparison.OrdinalIgnoreCase))
                {
                    font.Outline = true;
                }
                else if (nodeName.Equals("shadow", StringComparison.OrdinalIgnoreCase))
                {
                    font.Shadow = true;
                }
                else if (nodeName.Equals("condense", StringComparison.OrdinalIgnoreCase))
                {
                    font.Condense = true;
                }
                else if (nodeName.Equals("extend", StringComparison.OrdinalIgnoreCase))
                {
                    font.Extend = true;
                }
                else if (nodeName.Equals("color", StringComparison.OrdinalIgnoreCase))
                {
                    font.ColorValue = ParseColor(rPrReader);
                }
                else if (nodeName.Equals("sz", StringComparison.OrdinalIgnoreCase))
                {
                    string val = rPrReader.GetAttribute("val");
                    if (!string.IsNullOrEmpty(val))
                        font.Size = ParserUtils.ParseFloat(val);
                }
                else if (nodeName.Equals("u", StringComparison.OrdinalIgnoreCase))
                {
                    string val = rPrReader.GetAttribute("val");
                    font.Underline = string.IsNullOrEmpty(val) ? Font.UnderlineValue.Single : ParseUnderlineValue(val);
                }
                else if (nodeName.Equals("vertAlign", StringComparison.OrdinalIgnoreCase))
                {
                    string val = rPrReader.GetAttribute("val");
                    if (!string.IsNullOrEmpty(val))
                        font.VerticalAlign = ParseVerticalAlignValue(val);
                }
                else if (nodeName.Equals("scheme", StringComparison.OrdinalIgnoreCase))
                {
                    string val = rPrReader.GetAttribute("val");
                    if (!string.IsNullOrEmpty(val))
                        font.Scheme = ParseSchemeValue(val);
                }
            }

            return font;
        }

        /// <summary>
        /// Parses a color element and creates a Color object. Reader must be positioned on the &lt;color&gt; element.
        /// </summary>
        /// <param name="reader">XmlReader positioned on the color element</param>
        /// <returns>Color object, or null if no recognized color attribute is present</returns>
        private Color ParseColor(XmlReader reader)
        {
            string autoValue = reader.GetAttribute("auto");
            string indexedValue = reader.GetAttribute("indexed");
            string rgbValue = reader.GetAttribute("rgb");
            string themeValue = reader.GetAttribute("theme");
            string systemValue = reader.GetAttribute("system");
            string tintValue = reader.GetAttribute("tint");

            Color color = null;

            if (!string.IsNullOrEmpty(autoValue))
                color = Color.CreateAuto();
            else if (!string.IsNullOrEmpty(indexedValue))
                color = Color.CreateIndexed(ParserUtils.ParseInt(indexedValue));
            else if (!string.IsNullOrEmpty(rgbValue))
                color = Color.CreateRgb(rgbValue);
            else if (!string.IsNullOrEmpty(themeValue))
                color = Color.CreateTheme(ParserUtils.ParseInt(themeValue));
            else if (!string.IsNullOrEmpty(systemValue))
                color = Color.CreateSystem(SystemColor.MapStringToValue(systemValue));

            if (color != null && !string.IsNullOrEmpty(tintValue))
                color.Tint = ParserUtils.ParseFloat(tintValue);

            return color;
        }

        /// <summary>
        /// Processes a phonetic run (&lt;rPh&gt;) subtree
        /// </summary>
        /// <param name="rPhReader">XmlReader subtree positioned at the &lt;rPh&gt; open tag</param>
        /// <param name="formattedText">FormattedText to add the phonetic run to</param>
        private void ProcessPhoneticRun(XmlReader rPhReader, FormattedText formattedText)
        {
            string startBase = rPhReader.GetAttribute("sb");
            string endBase = rPhReader.GetAttribute("eb");
            string text = null;

            while (rPhReader.Read())
            {
                if (rPhReader.NodeType != XmlNodeType.Element)
                    continue;

                if (XmlStreamUtils.IsElement(rPhReader, "t"))
                {
                    using (XmlReader tReader = rPhReader.ReadSubtree())
                    {
                        tReader.Read(); // position at <t>
                        text = tReader.ReadElementContentAsString();
                    }
                }
            }

            if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(startBase) && !string.IsNullOrEmpty(endBase))
            {
                uint sb = (uint)ParserUtils.ParseInt(startBase);
                uint eb = (uint)ParserUtils.ParseInt(endBase);
                formattedText.AddPhoneticRun(text, sb, eb);

                if (capturePhoneticCharacters)
                {
                    phoneticsInfo.Add(new PhoneticInfo(text, startBase, endBase));
                }
            }
        }

        /// <summary>
        /// Processes phonetic properties (&lt;phoneticPr&gt;) from the current reader position
        /// </summary>
        /// <param name="reader">XmlReader positioned on the &lt;phoneticPr&gt; element</param>
        /// <param name="formattedText">FormattedText to set the properties on</param>
        private void ProcessPhoneticProperties(XmlReader reader, FormattedText formattedText)
        {
            string typeValue = reader.GetAttribute("type");
            string alignmentValue = reader.GetAttribute("alignment");

            PhoneticRun.PhoneticType type = PhoneticRun.PhoneticType.FullwidthKatakana;
            if (!string.IsNullOrEmpty(typeValue))
                type = ParsePhoneticType(typeValue);

            PhoneticRun.PhoneticAlignment alignment = PhoneticRun.PhoneticAlignment.Left;
            if (!string.IsNullOrEmpty(alignmentValue))
                alignment = ParsePhoneticAlignment(alignmentValue);

            formattedText.SetPhoneticProperties(new Font(), type, alignment);
        }

        /// <summary>
        /// Function to add determined phonetic tokens
        /// </summary>
        /// <param name="sb">Original StringBuilder</param>
        /// <returns>Text with added phonetic characters (after particular characters, in brackets)</returns>
        private string ProcessPhoneticCharacters(StringBuilder sb)
        {
            string text = sb.ToString();
            StringBuilder sb2 = new StringBuilder();
            int currentTextIndex = 0;
            foreach (PhoneticInfo info in phoneticsInfo)
            {
                sb2.Append(text.Substring(currentTextIndex, info.StartIndex + info.Length - currentTextIndex));
                sb2.Append('(').Append(info.Value).Append(')');
                currentTextIndex = info.StartIndex + info.Length;
            }
            sb2.Append(text.Substring(currentTextIndex));

            return sb2.ToString();
        }

        /// <summary>
        /// Parses underline value string to enum
        /// </summary>
        private Font.UnderlineValue ParseUnderlineValue(string value)
        {
            switch (value.ToLowerInvariant())
            {
                case "double":
                    return Font.UnderlineValue.Double;
                case "singleaccounting":
                    return Font.UnderlineValue.SingleAccounting;
                case "doubleaccounting":
                    return Font.UnderlineValue.DoubleAccounting;
                default:
                    return Font.UnderlineValue.Single;
            }
        }

        /// <summary>
        /// Parses vertical align value string to enum
        /// </summary>
        private Font.VerticalTextAlignValue ParseVerticalAlignValue(string value)
        {
            switch (value.ToLowerInvariant())
            {
                case "superscript":
                    return Font.VerticalTextAlignValue.Superscript;
                case "subscript":
                    return Font.VerticalTextAlignValue.Subscript;
                default:
                    return Font.VerticalTextAlignValue.Baseline;
            }
        }

        /// <summary>
        /// Parses scheme value string to enum
        /// </summary>
        private Font.SchemeValue ParseSchemeValue(string value)
        {
            switch (value.ToLowerInvariant())
            {
                case "major":
                    return Font.SchemeValue.Major;
                case "minor":
                    return Font.SchemeValue.Minor;
                default:
                    return Font.SchemeValue.None;
            }
        }

        /// <summary>
        /// Parses phonetic type value string to enum
        /// </summary>
        private PhoneticRun.PhoneticType ParsePhoneticType(string value)
        {
            switch (value.ToLowerInvariant())
            {
                case "halfwidthkatakana":
                    return PhoneticRun.PhoneticType.HalfwidthKatakana;
                case "hiragana":
                    return PhoneticRun.PhoneticType.Hiragana;
                case "noconversion":
                    return PhoneticRun.PhoneticType.NoConversion;
                default:
                    return PhoneticRun.PhoneticType.FullwidthKatakana;
            }
        }

        /// <summary>
        /// Parses phonetic alignment value string to enum
        /// </summary>
        private PhoneticRun.PhoneticAlignment ParsePhoneticAlignment(string value)
        {
            switch (value.ToLowerInvariant())
            {
                case "nocontrol":
                    return PhoneticRun.PhoneticAlignment.NoControl;
                case "center":
                    return PhoneticRun.PhoneticAlignment.Center;
                case "distributed":
                    return PhoneticRun.PhoneticAlignment.Distributed;
                default:
                    return PhoneticRun.PhoneticAlignment.Left;
            }
        }

        #endregion

        #region sub-classes
        /// <summary>
        /// Class to represent a phonetic transcription of character sequence.
        /// Note: Invalid values will lead to a crash. The specifications requires a start index, an end index and a value
        /// </summary>
        sealed class PhoneticInfo
        {
            /// <summary>
            /// Transcription value
            /// </summary>
            public string Value { get; private set; }
            /// <summary>
            /// Absolute start index within the original string
            /// </summary>
            public int StartIndex { get; private set; }
            /// <summary>
            /// Number of characters of the original string that are described by this transcription token
            /// </summary>
            public int Length { get; private set; }

            /// <summary>
            /// Constructor with parameters
            /// </summary>
            /// <param name="value">Transcription value</param>
            /// <param name="start">Absolute start index as string</param>
            /// <param name="end">Absolute end index as string</param>
            public PhoneticInfo(string value, string start, string end)
            {
                Value = value;
                StartIndex = ParserUtils.ParseInt(start);
                Length = ParserUtils.ParseInt(end) - StartIndex;

            }
        }
        #endregion
    }
}
