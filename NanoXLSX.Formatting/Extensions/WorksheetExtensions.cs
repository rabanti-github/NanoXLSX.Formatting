
using System.ComponentModel;
using NanoXLSX.Exceptions;


namespace NanoXLSX
{
    /// <summary>
    /// Writer extension methods for the <see cref="NanoXLSX.Worksheet">Workbook</see> class
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public static class WorksheetExtensions
    {
        public static void AddCell(this Worksheet worksheet, FormattedText formattedText, int columnNumber, int rowNumber)
        {
            //AddAuxiliaryData(worksheet, formattedText, columnNumber, rowNumber);
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
    }
}
