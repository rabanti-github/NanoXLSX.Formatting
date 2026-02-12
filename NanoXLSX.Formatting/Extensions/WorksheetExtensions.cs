/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using System.ComponentModel;
using NanoXLSX.Exceptions;
using NanoXLSX.Styles;

namespace NanoXLSX
{
    /// <summary>
    /// Writer extension methods for the <see cref="NanoXLSX.Worksheet">Worksheet</see> class
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
        public static void AddFormattedTextCell(this Worksheet worksheet, FormattedText formattedText, int columnNumber, int rowNumber)
        {
            AddFormattedText(worksheet, formattedText, columnNumber, rowNumber);
        }

        /// <summary>
        /// Adds a formatted text to the specified cell with a style. If the WrapText property of the formatted text is set to true, the LineBreakStyle will be merged with the provided style automatically.
        /// </summary>
        /// <param name="worksheet">Worksheet reference</param>
        /// <param name="formattedText">Formatted text object</param>
        /// <param name="columnNumber">Column number (zero based)</param>
        /// <param name="rowNumber">Row number (zero based)</param>
        /// <param name="style">Style to apply to the cell. Will be merged with <see cref="FormattedText.LineBreakStyle"/> if <see cref="FormattedText.WrapText"/> true</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the passed text object is null</exception>
        /// <exception cref="RangeException">Throws a RangeException if the passed cell address is out of range</exception>

        public static void AddFormattedTextCell(this Worksheet worksheet, FormattedText formattedText, int columnNumber, int rowNumber, Style style)
        {
            AddFormattedText(worksheet, formattedText, columnNumber, rowNumber, style);
        }

        /// <summary>
        /// Adds a formatted text to the specified cell address. If the WrapText property of the formatted text is set to true, the LineBreakStyle will be applied automatically.
        /// </summary>
        /// <param name="worksheet">Worksheet reference</param>
        /// <param name="formattedText">Formatted text object</param>
        /// <param name="address">Cell address in the format A1 - XFD1048576 (e.g., "A1", "$B$5", "c12"). Case-insensitive, dollar signs are ignored</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the passed text object is null</exception>
        /// <exception cref="FormatException">Throws a FormatException if the passed address is malformed</exception>
        /// <exception cref="RangeException">Throws a RangeException if the passed cell address is out of range</exception>
        public static void AddFormattedTextCell(this Worksheet worksheet, FormattedText formattedText, string address)
        {
            int column;
            int row;
            Cell.ResolveCellCoordinate(address, out column, out row);
            AddFormattedText(worksheet, formattedText, column, row);
        }

        /// <summary>
        /// Adds a formatted text to the specified cell address with a style. If the WrapText property of the formatted text is set to true, the LineBreakStyle will be merged with the provided style automatically.
        /// </summary>
        /// <param name="worksheet">Worksheet reference</param>
        /// <param name="formattedText">Formatted text object</param>
        /// <param name="address">Cell address in the format A1 - XFD1048576 (e.g., "A1", "$B$5", "c12"). Case-insensitive, dollar signs are ignored</param>
        /// <param name="style">Style to apply to the cell. Will be merged with <see cref="FormattedText.LineBreakStyle"/> if <see cref="FormattedText.WrapText"/> true</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the passed text object is null</exception>
        /// <exception cref="FormatException">Throws a FormatException if the passed address is malformed</exception>
        /// <exception cref="RangeException">Throws a RangeException if the passed cell address is out of range</exception>
        public static void AddFormattedTextCell(this Worksheet worksheet, FormattedText formattedText, string address, Style style)
        {
            int column;
            int row;
            Cell.ResolveCellCoordinate(address, out column, out row);
            AddFormattedText(worksheet, formattedText, column, row, style);
        }

        /// <summary>
        /// Adds a formatted text to the next cell position. The direction of the next cell depends on the current cell direction (default is <see cref="Worksheet.CellDirection.ColumnToColumn"/>). 
        /// If the WrapText property of the formatted text is set to true, the LineBreakStyle will be applied automatically.
        /// </summary>
        /// <param name="worksheet">Worksheet reference</param>
        /// <param name="formattedText">Formatted text object</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the passed text object is null</exception>
        /// <exception cref="RangeException">Throws a RangeException if the next cell address is out of range</exception>
        public static void AddNextFormattedTextCell(this Worksheet worksheet, FormattedText formattedText)
        {
            AddFormattedText(worksheet, formattedText);
        }

        /// <summary>
        /// Adds a formatted text to the next cell position with a style. The direction of the next cell depends on the current cell direction (default is <see cref="Worksheet.CellDirection.ColumnToColumn"/>). 
        /// If the WrapText property of the formatted text is set to true, the LineBreakStyle will be merged with the provided style automatically.
        /// </summary>
        /// <param name="worksheet">Worksheet reference</param>
        /// <param name="formattedText">Formatted text object</param>
        /// <param name="style">Style to apply to the cell. Will be merged with <see cref="FormattedText.LineBreakStyle"/> if <see cref="FormattedText.WrapText"/> is true</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the passed text object is null</exception>
        /// <exception cref="RangeException">Throws a RangeException if the next cell address is out of range</exception>
        public static void AddNextFormattedTextCell(this Worksheet worksheet, FormattedText formattedText, Style style)
        {
            AddFormattedText(worksheet, formattedText, style);
        }

        /// <summary>
        /// Adds a list of formatted texts to a defined cell range. Values are distributed column-by-column from the start address. 
        /// Null values will be inserted as empty cells. If the WrapText property of a formatted text is set to true, the LineBreakStyle will be applied automatically to that cell.
        /// </summary>
        /// <param name="worksheet">Worksheet reference</param>
        /// <param name="values">List of formatted text objects to insert. Null values are treated as empty cells</param>
        /// <param name="startAddress">Start address of the range as string</param>
        /// <param name="endAddress">End address of the range as string</param>
        /// <param name="style">Optional style to apply to all cells in the range. Will be merged with <see cref="FormattedText.LineBreakStyle"/> for cells where <see cref="FormattedText.WrapText"/> is true</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the list of formatted texts is null or empty</exception>
        /// <exception cref="RangeException">Throws a RangeException if the number of values does not match the number of cells in the range, or if any address is out of range</exception>
        public static void AddFormattedTextCellRange(this Worksheet worksheet, IReadOnlyList<FormattedText> values, string startAddress, string endAddress, Style style = null)
        {
            AddFormattedTextRange(worksheet, values, (Address)startAddress, (Address)endAddress, style);
        }

        /// <summary>
        /// Adds a list of formatted texts to a defined cell range with a style applied to all cells. Values are distributed column-by-column from the start address. 
        /// Null values will be inserted as empty cells with the specified style. If the WrapText property of a formatted text is set to true, the LineBreakStyle will be merged with the provided style for that cell.
        /// </summary>
        /// <param name="worksheet">Worksheet reference</param>
        /// <param name="values">List of formatted text objects to insert. Null values are treated as empty cells</param>
        /// <param name="startAddress">Start address of the range</param>
        /// <param name="endAddress">End address of the range</param>
        /// <param name="style">Optional style to apply to all cells in the range. Will be merged with <see cref="FormattedText.LineBreakStyle"/> for cells where <see cref="FormattedText.WrapText"/> is true</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the list of formatted texts is null or empty</exception>
        /// <exception cref="RangeException">Throws a RangeException if the number of values does not match the number of cells in the range, or if any address is out of range</exception>
        public static void AddFormattedTextCellRange(this Worksheet worksheet, IReadOnlyList<FormattedText> values, Address startAddress, Address endAddress, Style style = null)
        {
            AddFormattedTextRange(worksheet, values, startAddress, endAddress, style);
        }

        /// <summary>
        /// Adds a list of formatted texts to a defined cell range with a style applied to all cells. Values are distributed column-by-column from the start of the range. 
        /// Null values will be inserted as empty cells with the specified style. If the WrapText property of a formatted text is set to true, the LineBreakStyle will be merged with the provided style for that cell.
        /// </summary>
        /// <param name="worksheet">Worksheet reference</param>
        /// <param name="values">List of formatted text objects to insert. Null values are treated as empty cells</param>
        /// <param name="cellRange">Cell range in the format A1:B10 (e.g., "A1:C5", "$A$1:$C$5", "a1:c5"). Case-insensitive, dollar signs are ignored</param>
        /// <param name="style">Optional style to apply to all cells in the range. Will be merged with <see cref="FormattedText.LineBreakStyle"/> for cells where <see cref="FormattedText.WrapText"/> is true</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the list of formatted texts is null or empty</exception>
        /// <exception cref="FormatException">Throws a FormatException if the cell range format is malformed</exception>
        /// <exception cref="RangeException">Throws a RangeException if the number of values does not match the number of cells in the range, or if any address is out of range</exception>
        public static void AddFormattedTextCellRange(this Worksheet worksheet, IReadOnlyList<FormattedText> values, string cellRange, Style style = null)
        {
            Range range = new Range(cellRange);
            AddFormattedTextRange(worksheet, values, range.StartAddress, range.EndAddress, style);
        }


        /// <summary>
        /// Adds a formatted text to the next cell position with an optional style.
        /// </summary>
        /// <param name="worksheet">Worksheet reference</param>
        /// <param name="formattedText">Formatted text object</param>
        /// <param name="columnNumber">Column number (zero based)</param>
        /// <param name="rowNumber">Row number (zero based)</param>
        /// <param name="cellStyle">Style to apply to the cell. Will be merged with <see cref="FormattedText.LineBreakStyle"/> if <see cref="FormattedText.WrapText"/> is true</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the list of formatted texts is null or empty</exception>
        private static void AddFormattedText(Worksheet worksheet, FormattedText formattedText, int columnNumber, int rowNumber, Style cellStyle = null)
        {
            if (formattedText == null)
            {
                throw new WorksheetException("A formatted text to add cannot be null");
            }
            if (formattedText.WrapText)
            {
                if (cellStyle == null)
                {
                    worksheet.AddCell(formattedText, columnNumber, rowNumber, FormattedText.LineBreakStyle);
                }
                else
                {
                    Style mergedStyle = cellStyle.Append(FormattedText.LineBreakStyle);
                    worksheet.AddCell(formattedText, columnNumber, rowNumber, mergedStyle);
                }
            }
            else
            {
                worksheet.AddCell(formattedText, columnNumber, rowNumber, cellStyle);
            }
        }

        /// <summary>
        /// Adds a formatted text to the next cell position with an optional style.
        /// </summary>
        /// <param name="worksheet">Worksheet reference</param>
        /// <param name="formattedText">Formatted text object</param>
        /// <param name="cellStyle">Style to apply to all cells in the range. Will be merged with <see cref="FormattedText.LineBreakStyle"/> for cells where WrapText is true</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the list of formatted texts is null or empty</exception>
        private static void AddFormattedText(Worksheet worksheet, FormattedText formattedText, Style cellStyle = null)
        {
            if (formattedText == null)
            {
                throw new WorksheetException("A formatted text to add cannot be null");
            }
            if (formattedText.WrapText)
            {
                if (cellStyle == null)
                {
                    worksheet.AddNextCell(formattedText, FormattedText.LineBreakStyle);
                }
                else
                {
                    Style mergedStyle = cellStyle.Append(FormattedText.LineBreakStyle);
                    worksheet.AddNextCell(formattedText, mergedStyle);
                }
            }
            else
            {
                worksheet.AddNextCell(formattedText, cellStyle);
            }
        }

        /// <summary>
        /// Adds a list of formatted texts to a defined cell range with an optional style applied to all cells.
        /// </summary>
        /// <param name="worksheet">Worksheet reference</param>
        /// <param name="formattedTexts">Formatted text objects</param>
        /// <param name="startAddress">Start address of the range</param>
        /// <param name="endAddress">End address of the range</param>
        /// <param name="cellStyle">Style to apply to all cells in the range. Will be merged with <see cref="FormattedText.LineBreakStyle"/> for cells where <see cref="FormattedText.WrapText"/> is true</param>
        /// <exception cref="WorksheetException">Throws a WorksheetException if the list of formatted texts is null or empty</exception>
        private static void AddFormattedTextRange(Worksheet worksheet, IReadOnlyList<FormattedText> formattedTexts, Address startAddress, Address endAddress, Style cellStyle = null)
        {
            if (formattedTexts == null || formattedTexts.Count == 0)
            {
                throw new WorksheetException("A range of formatted texts to add cannot be null or empty");
            }
            List<Cell> cells = new List<Cell>(formattedTexts.Count);
            foreach (FormattedText text in formattedTexts)
            {
                if (text == null || string.IsNullOrEmpty(text.PlainText))
                {
                    Cell emptyCell = new Cell(null, Cell.CellType.Empty);
                    if (cellStyle != null)
                    {
                        emptyCell.SetStyle(cellStyle);
                    }
                    cells.Add(emptyCell);
                    continue;
                }
                Cell cell = new Cell(text, Cell.CellType.String);
                if (text.WrapText && cellStyle != null)
                {
                    Style merged = cellStyle.Append(FormattedText.LineBreakStyle);
                    cell.SetStyle(merged);
                }
                else if (text.WrapText && cellStyle == null)
                {
                    cell.SetStyle(FormattedText.LineBreakStyle);
                }
                else if (!text.WrapText && cellStyle != null)
                {
                    cell.SetStyle(cellStyle);
                }
                cells.Add(cell);
            }
            if (cellStyle != null)
            {
                worksheet.AddCellRange(cells, startAddress, endAddress, cellStyle);
            }
            else
            {
                worksheet.AddCellRange(cells, startAddress, endAddress);
            }
        }

    }
}
