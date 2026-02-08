/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Styles;
using NanoXLSX.Utils.Xml;

namespace NanoXLSX.Internal
{
    /// <summary>
    /// Class to resolve font IDs for phonetic properties in formatted text cells.
    /// </summary>
    [NanoXlsxQueuePlugIn(PlugInUUID = "FORMATTED_TEXT_FONT_ID_RESOLVER", QueueUUID = PlugInUUID.WriterPrependingQueue)]
    internal class FontIdResolver : IPluginWriter
    {
        #region privateFields
        private StyleManager styleManager;
        #endregion

        #region properties
        /// <summary>
        /// Current workbook
        /// </summary>
        public Workbook Workbook { get; set; }

        /// <summary>
        /// XML element representation (interface implementation). Not used in this class.
        /// </summary>
        [ExcludeFromCodeCoverage] // NoOp
        public XmlElement XmlElement { get { return null; } } // NoOp
        #endregion

        #region methods
        /// <summary>
        /// Execute method (interface implementation).
        /// </summary>
        public void Execute()
        {
            Font[] fonts = styleManager.GetFonts();
            Dictionary<Font, int> fontIndexLookup = fonts
                .Select((font, index) => new { font, index })
                .ToDictionary(x => x.font, x => x.index);

            for (int i = 0; i < Workbook.Worksheets.Count; i++)
            {
                List<FormattedText> formattedTextCells = Workbook.Worksheets[i].Cells.Values
                    .Select(c => c.Value)
                    .OfType<FormattedText>()
                    .Where(ft => ft.PhoneticProperties?.FontReference != null)
                    .ToList();
                if (formattedTextCells.Count == 0)
                {
                    continue;
                }
                foreach (FormattedText formattedText in formattedTextCells)
                {
                    Extensions.PhoneticProperties props = formattedText.PhoneticProperties;
                    Font fontReference = props.FontReference;
                    if (fontIndexLookup.TryGetValue(fontReference, out int fontIndex))
                    {
                        props.FontId = fontIndex;
                    }
                }
            }
        }

        /// <summary>
        /// Initialization method (interface implementation).
        /// </summary>
        /// <param name="baseWriter">Underlaying base writer instance</param>
        public void Init(IBaseWriter baseWriter)
        {
            this.Workbook = baseWriter.Workbook;
            this.styleManager = baseWriter.Styles;
        }
        #endregion
    }
}
