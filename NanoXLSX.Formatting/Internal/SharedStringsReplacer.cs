/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.IO;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;

namespace NanoXLSX.Internal.Readers
{
    [NanoXlsxQueuePlugIn(PlugInUUID = "FORMATTED_TEXT_SHARED_STRINGS_REPLACER", QueueUUID = PlugInUUID.ReaderAppendingQueue)]
    internal class SharedStringsReplacer : IPluginQueueReader
    {
        #region properties
        /// <summary>
        /// Reader options
        /// </summary>
        public IOptions Options { get; set; }
        /// <summary>
        /// Current workbook
        /// </summary>
        public Workbook Workbook { get; set; }
        /// <summary>
        /// Reference to a ReaderPlugInHandler, to be used for post operations in the <see cref="Execute"/> method
        /// </summary>
        public Action<MemoryStream, Workbook, string, IOptions, int?> InlinePluginHandler { get; set; }
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
            this.Workbook = workbook;
            this.Options = readerOptions;
            this.InlinePluginHandler = inlinePluginHandler;
        }

        /// <summary>
        /// Execution method (interface implementation)
        /// </summary>
        public void Execute()
        {
            Dictionary<string, FormattedText> formattedTexts = Workbook.AuxiliaryData.GetData<Dictionary<string, FormattedText>>(PlugInUUID.SharedStringsReader, FormattedSharedStringsReader.AUXILIARY_DATA_ID);
            if (formattedTexts == null)
            {
                return;
            }
            for (int i = 0; i < Workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = Workbook.Worksheets[i];
                foreach (KeyValuePair<string, FormattedText> entry in formattedTexts)
                {
                    List<Cell> cells = sheet.CellsByValue(entry.Key);
                    if (cells != null && cells.Count > 0)
                    {
                        foreach (Cell cell in cells)
                        {
                            sheet.Cells[cell.CellAddress].Value = entry.Value;
                        }
                    }
                }
            }
        }
        #endregion
    }
}
