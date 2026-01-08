/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.ComponentModel;
using NanoXLSX.Exceptions;
using NanoXLSX.Styles;

namespace NanoXLSX
{
    /// <summary>
    /// Writer extension methods for the <see cref="NanoXLSX.Worksheet">Workbook</see> class
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public static class WorksheetExtensions
    {
        /// <summary>
        /// Adds a formatted text to the specified cell. If the WrapText property of the formatted text is set to true, the LineBreakStyle will be applied automatically.
        /// </summary>
        /// <param name="worksheet">Worksheet reference</param>
        /// <param name="formattedText">Formatted text object</param>
        /// <param name="columnNumber">Column number (zero based)</param>
        /// <param name="rowNumber">Row number (zero based)</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the passed text object is null</exception>
        /// <exception cref="RangeException">Throws a RangeException if the passed cell address is out of range</exception>
        public static void AddCell(this Worksheet worksheet, FormattedText formattedText, int columnNumber, int rowNumber)
        {
            AddFormattedText(worksheet, formattedText, columnNumber, rowNumber);
        }

        public static void Addcell(this Worksheet worksheet, FormattedText formattedText, string address)
        {
            int column;
            int row;
            Cell.ResolveCellCoordinate(address, out column, out row);
            AddFormattedText(worksheet, formattedText, column, row);
        }

        public static void AddNextCell(this Worksheet worksheet, FormattedText formattedText)
        {
            worksheet.AddFormattedText(formattedText);
        }

        public static void AddNextCell(this Worksheet worksheet, bool incremental, FormattedText formattedText)
        {
            worksheet.AddFormattedText(formattedText, incremental);
        }

        private static void AddFormattedText(Worksheet worksheet, FormattedText formattedText, int columnNumber, int rowNumber)
        {
            if (formattedText == null)
            {
                throw new WorksheetException("A formatted text to add cannot be null");
            }
            if (formattedText.WrapText)
            {
                worksheet.AddCell(formattedText, columnNumber, rowNumber, FormattedText.LineBreakStyle);
            }
            else
            {
                worksheet.AddCell(formattedText, columnNumber, rowNumber);
            }
        }

        private static void AddFormattedText(this Worksheet worksheet, FormattedText formattedText, bool incemental = false)
        {
            if (formattedText == null)
            {
                throw new WorksheetException("A formatted text to add cannot be null");
            }
            if (formattedText.WrapText)
            {
                Cell cell = new Cell(formattedText, Cell.CellType.Default);
                worksheet.AddNextCell(cell, incemental, new Style());
            }
            else
            {
                worksheet.AddNextCell(formattedText);
            }
        }

    }
}
