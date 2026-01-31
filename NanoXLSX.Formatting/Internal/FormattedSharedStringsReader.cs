/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
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
        private MemoryStream stream;
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
        public Action<MemoryStream, Workbook, string, IOptions, int?> InlinePluginHandler { get; set; }
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
        /// <param name="stream">MemoryStream to be read</param>
        /// <param name="workbook">Workbook reference</param>
        /// <param name="readerOptions">Reader options</param>
        /// <param name="inlinePluginHandler">Inline plug-in handler</param>
        public void Init(MemoryStream stream, Workbook workbook, IOptions readerOptions, Action<MemoryStream, Workbook, string, IOptions, int?> inlinePluginHandler)
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
                    XmlDocument xr = new XmlDocument
                    {
                        XmlResolver = null
                    };
                    bool hasFormattedText = false;
                    using (XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings() { XmlResolver = null }))
                    {
                        xr.Load(reader);
                        StringBuilder sb = new StringBuilder();
                        foreach (XmlNode node in xr.DocumentElement.ChildNodes)
                        {
                            if (node.LocalName.Equals("si", StringComparison.OrdinalIgnoreCase))
                            {
                                sb.Clear();
                                phoneticsInfo.Clear();

                                FormattedText formattedText = ProcessSharedStringItem(node, ref sb);
                                string textValue;
                                if (capturePhoneticCharacters)
                                {
                                    textValue = ProcessPhoneticCharacters(sb);
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
                    }
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
        /// Processes a shared string item (&lt;si&gt;) node and creates FormattedText if it contains runs
        /// </summary>
        /// <param name="siNode">The si node</param>
        /// <param name="sb">StringBuilder for plain text extraction</param>
        /// <returns>FormattedText if the item has formatting, null otherwise</returns>
        private FormattedText ProcessSharedStringItem(XmlNode siNode, ref StringBuilder sb)
        {
            bool hasRuns = false;
            bool hasPhoneticRuns = false;
            XmlNode phoneticPropertiesNode = null;

            // Check if this is a formatted text entry
            foreach (XmlNode childNode in siNode.ChildNodes)
            {
                if (childNode.LocalName.Equals("r", StringComparison.OrdinalIgnoreCase))
                {
                    hasRuns = true;
                }
                else if (childNode.LocalName.Equals("rPh", StringComparison.OrdinalIgnoreCase))
                {
                    hasPhoneticRuns = true;
                }
                else if (childNode.LocalName.Equals("phoneticPr", StringComparison.OrdinalIgnoreCase))
                {
                    phoneticPropertiesNode = childNode;
                }
            }

            if (!hasRuns && !hasPhoneticRuns && phoneticPropertiesNode == null)
            {
                // Simple text node, just extract plain text
                GetTextToken(siNode, ref sb);
                return null;
            }

            // Create FormattedText object
            FormattedText formattedText = new FormattedText();

            // Process text runs
            if (hasRuns)
            {
                foreach (XmlNode childNode in siNode.ChildNodes)
                {
                    if (childNode.LocalName.Equals("r", StringComparison.OrdinalIgnoreCase))
                    {
                        ProcessTextRun(childNode, formattedText, ref sb);
                    }
                }
            }
            else
            {
                // No runs but has phonetic info - extract text as single run
                GetTextToken(siNode, ref sb);
            }

            // Process phonetic runs
            if (hasPhoneticRuns)
            {
                foreach (XmlNode childNode in siNode.ChildNodes)
                {
                    if (childNode.LocalName.Equals("rPh", StringComparison.OrdinalIgnoreCase))
                    {
                        ProcessPhoneticRun(childNode, formattedText);
                    }
                }
            }

            // Process phonetic properties
            if (phoneticPropertiesNode != null)
            {
                ProcessPhoneticProperties(phoneticPropertiesNode, formattedText);
            }

            return formattedText;
        }

        /// <summary>
        /// Processes a text run (&lt;r&gt;) node
        /// </summary>
        /// <param name="runNode">The run node</param>
        /// <param name="formattedText">FormattedText to add the run to</param>
        /// <param name="sb">StringBuilder for plain text extraction</param>
        private void ProcessTextRun(XmlNode runNode, FormattedText formattedText, ref StringBuilder sb)
        {
            Font fontStyle = null;
            string text = null;

            foreach (XmlNode childNode in runNode.ChildNodes)
            {
                if (childNode.LocalName.Equals("rPr", StringComparison.OrdinalIgnoreCase))
                {
                    fontStyle = ParseRunProperties(childNode);
                }
                else if (childNode.LocalName.Equals("t", StringComparison.OrdinalIgnoreCase))
                {
                    text = childNode.InnerText;
                    sb.Append(text);
                }
            }

            if (!string.IsNullOrEmpty(text))
            {
                formattedText.AddRun(text, fontStyle);
            }
        }

        /// <summary>
        /// Parses run properties (&lt;rPr&gt;) and creates a Font object
        /// </summary>
        /// <param name="rPrNode">The rPr node</param>
        /// <returns>Font object with parsed properties</returns>
        private Font ParseRunProperties(XmlNode rPrNode)
        {
            Font font = new Font();

            foreach (XmlNode childNode in rPrNode.ChildNodes)
            {
                string nodeName = childNode.LocalName;

                if (nodeName.Equals("rFont", StringComparison.OrdinalIgnoreCase))
                {
                    font.Name = GetAttributeValue(childNode, "val");
                }
                else if (nodeName.Equals("charset", StringComparison.OrdinalIgnoreCase))
                {
                    string charsetValue = GetAttributeValue(childNode, "val");
                    if (!string.IsNullOrEmpty(charsetValue))
                    {
                        font.Charset = (Font.CharsetValue)ParserUtils.ParseInt(charsetValue);
                    }
                }
                else if (nodeName.Equals("family", StringComparison.OrdinalIgnoreCase))
                {
                    string familyValue = GetAttributeValue(childNode, "val");
                    if (!string.IsNullOrEmpty(familyValue))
                    {
                        font.Family = (Font.FontFamilyValue)ParserUtils.ParseInt(familyValue);
                    }
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
                    font.ColorValue = ParseColor(childNode);
                }
                else if (nodeName.Equals("sz", StringComparison.OrdinalIgnoreCase))
                {
                    string sizeValue = GetAttributeValue(childNode, "val");
                    if (!string.IsNullOrEmpty(sizeValue))
                    {
                        font.Size = ParserUtils.ParseFloat(sizeValue);
                    }
                }
                else if (nodeName.Equals("u", StringComparison.OrdinalIgnoreCase))
                {
                    string underlineValue = GetAttributeValue(childNode, "val");
                    if (string.IsNullOrEmpty(underlineValue))
                    {
                        font.Underline = Font.UnderlineValue.Single;
                    }
                    else
                    {
                        font.Underline = ParseUnderlineValue(underlineValue);
                    }
                }
                else if (nodeName.Equals("vertAlign", StringComparison.OrdinalIgnoreCase))
                {
                    string vertAlignValue = GetAttributeValue(childNode, "val");
                    if (!string.IsNullOrEmpty(vertAlignValue))
                    {
                        font.VerticalAlign = ParseVerticalAlignValue(vertAlignValue);
                    }
                }
                else if (nodeName.Equals("scheme", StringComparison.OrdinalIgnoreCase))
                {
                    string schemeValue = GetAttributeValue(childNode, "val");
                    if (!string.IsNullOrEmpty(schemeValue))
                    {
                        font.Scheme = ParseSchemeValue(schemeValue);
                    }
                }
            }

            return font;
        }

        /// <summary>
        /// Parses a color node and creates a Color object
        /// </summary>
        /// <param name="colorNode">The color node</param>
        /// <returns>Color object</returns>
        private Color ParseColor(XmlNode colorNode)
        {
            string autoValue = GetAttributeValue(colorNode, "auto");
            string indexedValue = GetAttributeValue(colorNode, "indexed");
            string rgbValue = GetAttributeValue(colorNode, "rgb");
            string themeValue = GetAttributeValue(colorNode, "theme");
            string systemValue = GetAttributeValue(colorNode, "system");
            string tintValue = GetAttributeValue(colorNode, "tint");

            Color color = null; //= new Color();

            if (!string.IsNullOrEmpty(autoValue))
            {
                color = Color.CreateAuto();
            }
            else if (!string.IsNullOrEmpty(indexedValue))
            {
                color = Color.CreateIndexed(ParserUtils.ParseInt(indexedValue));
            }
            else if (!string.IsNullOrEmpty(rgbValue))
            {
                color = Color.CreateRgb(rgbValue);
            }
            else if (!string.IsNullOrEmpty(themeValue))
            {
                color = Color.CreateTheme(ParserUtils.ParseInt(themeValue));
            }
            else if (!string.IsNullOrEmpty(systemValue))
            {
                color = Color.CreateSystem(SystemColor.MapStringToValue(systemValue));
            }
            if (color != null && !string.IsNullOrEmpty(tintValue))
            {
                color.Tint = ParserUtils.ParseFloat(tintValue);
            }

            return color;
        }

        /// <summary>
        /// Processes a phonetic run (&lt;rPh&gt;) node
        /// </summary>
        /// <param name="rPhNode">The rPh node</param>
        /// <param name="formattedText">FormattedText to add the phonetic run to</param>
        private void ProcessPhoneticRun(XmlNode rPhNode, FormattedText formattedText)
        {
            string startBase = GetAttributeValue(rPhNode, "sb");
            string endBase = GetAttributeValue(rPhNode, "eb");
            string text = null;

            foreach (XmlNode childNode in rPhNode.ChildNodes)
            {
                if (childNode.LocalName.Equals("t", StringComparison.OrdinalIgnoreCase))
                {
                    text = childNode.InnerText;
                    break;
                }
            }

            if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(startBase) && !string.IsNullOrEmpty(endBase))
            {
                uint sb = (uint)ParserUtils.ParseInt(startBase);
                uint eb = (uint)ParserUtils.ParseInt(endBase);
                formattedText.AddPhoneticRun(text, sb, eb);

                // Also capture for plain text processing
                if (capturePhoneticCharacters)
                {
                    phoneticsInfo.Add(new PhoneticInfo(text, startBase, endBase));
                }
            }
        }

        /// <summary>
        /// Processes phonetic properties (&lt;phoneticPr&gt;) node
        /// </summary>
        /// <param name="phoneticPrNode">The phoneticPr node</param>
        /// <param name="formattedText">FormattedText to set the properties on</param>
        private void ProcessPhoneticProperties(XmlNode phoneticPrNode, FormattedText formattedText)
        {
            string fontIdValue = GetAttributeValue(phoneticPrNode, "fontId");
            string typeValue = GetAttributeValue(phoneticPrNode, "type");
            string alignmentValue = GetAttributeValue(phoneticPrNode, "alignment");

            // Create a basic font reference for phonetic properties
            Font fontReference = new Font();
            if (!string.IsNullOrEmpty(fontIdValue))
            {
                // Font ID is stored but not resolved here
                // This is just a placeholder for the reference
            }

            PhoneticRun.PhoneticType type = PhoneticRun.PhoneticType.FullwidthKatakana;
            if (!string.IsNullOrEmpty(typeValue))
            {
                type = ParsePhoneticType(typeValue);
            }

            PhoneticRun.PhoneticAlignment alignment = PhoneticRun.PhoneticAlignment.Left;
            if (!string.IsNullOrEmpty(alignmentValue))
            {
                alignment = ParsePhoneticAlignment(alignmentValue);
            }

            formattedText.SetPhoneticProperties(fontReference, type, alignment);
        }

        /// <summary>
        /// Function collects text tokens recursively in case of a split by formatting
        /// </summary>
        /// <param name="node">Root node to process</param>
        /// <param name="sb">StringBuilder reference</param>
        private void GetTextToken(XmlNode node, ref StringBuilder sb)
        {
            if (node.LocalName.Equals("rPh", StringComparison.OrdinalIgnoreCase))
            {
                if (capturePhoneticCharacters && !string.IsNullOrEmpty(node.InnerText))
                {
                    string start = node.Attributes.GetNamedItem("sb").InnerText;
                    string end = node.Attributes.GetNamedItem("eb").InnerText;
                    phoneticsInfo.Add(new PhoneticInfo(node.InnerText, start, end));
                }
                return;
            }

            if (node.LocalName.Equals("t", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(node.InnerText))
            {
                sb.Append(node.InnerText);
            }
            if (node.HasChildNodes)
            {
                foreach (XmlNode childNode in node.ChildNodes)
                {
                    GetTextToken(childNode, ref sb);
                }
            }
        }

        /// <summary>
        /// Function to add determined phonetic tokens
        /// </summary>
        /// <param name="sb">Original StringBuilder</param>
        /// <returns>Text with added phonetic characters (after particular characters, in brackets)</returns>
        private string ProcessPhoneticCharacters(StringBuilder sb)
        {
            if (phoneticsInfo.Count == 0)
            {
                return sb.ToString();
            }
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
        /// Gets an attribute value from a node
        /// </summary>
        /// <param name="node">XML node</param>
        /// <param name="attributeName">Attribute name</param>
        /// <returns>Attribute value or null if not found</returns>
        private string GetAttributeValue(XmlNode node, string attributeName)
        {
            XmlNode attribute = node.Attributes?.GetNamedItem(attributeName);
            return attribute?.InnerText;
        }

        /// <summary>
        /// Parses underline value string to enum
        /// </summary>
        /// <param name="value">String value</param>
        /// <returns>UnderlineValue enum</returns>
        private Font.UnderlineValue ParseUnderlineValue(string value)
        {
            switch (value.ToLowerInvariant())
            {
                case "single":
                    return Font.UnderlineValue.Single;
                case "double":
                    return Font.UnderlineValue.Double;
                case "singleaccounting":
                    return Font.UnderlineValue.SingleAccounting;
                case "doubleaccounting":
                    return Font.UnderlineValue.DoubleAccounting;
                default:
                    return Font.UnderlineValue.None;
            }
        }

        /// <summary>
        /// Parses vertical align value string to enum
        /// </summary>
        /// <param name="value">String value</param>
        /// <returns>VerticalTextAlignValue enum</returns>
        private Font.VerticalTextAlignValue ParseVerticalAlignValue(string value)
        {
            switch (value.ToLowerInvariant())
            {
                case "baseline":
                    return Font.VerticalTextAlignValue.Baseline;
                case "superscript":
                    return Font.VerticalTextAlignValue.Superscript;
                case "subscript":
                    return Font.VerticalTextAlignValue.Subscript;
                default:
                    return Font.VerticalTextAlignValue.None;
            }
        }

        /// <summary>
        /// Parses scheme value string to enum
        /// </summary>
        /// <param name="value">String value</param>
        /// <returns>SchemeValue enum</returns>
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
        /// <param name="value">String value</param>
        /// <returns>PhoneticType enum</returns>
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
        /// <param name="value">String value</param>
        /// <returns>PhoneticAlignment enum</returns>
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
