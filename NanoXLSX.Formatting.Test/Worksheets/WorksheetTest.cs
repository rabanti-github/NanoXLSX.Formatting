using NanoXLSX.Extensions;
using NanoXLSX.Registry;
using NanoXLSX.Styles;
using Xunit;


namespace NanoXLSX.Formatting.Test.Worksheets
{
    public class WorksheetTest
    {
        [Fact]
        public void SampleTestMethod()
        {
            PlugInLoader.InjectPlugins(new System.Collections.Generic.List<System.Type>
            {
                 typeof(Internal.Readers.FormattedSharedStringsReader),
                 typeof(Internal.Readers.SharedStringsReplacer)
            });
            Workbook workbook = new Workbook("worksheet1");
            FormattedText formattedText = new FormattedText();
            InlineStyleBuilder builder = new InlineStyleBuilder();
            Font style1 = builder.FontName("Arial").Bold().ColorArgb("CC00FF").Build();
            Font style2 = builder.FontName("Calibri").Italic().ColorArgb("00FFFF").Shadow().Build();
            formattedText.AddRun("Line1", style1);
            formattedText.AddLineBreak();
            formattedText.AddRun(" Line2", style2);
            workbook.CurrentWorksheet.AddFormattedTextCell(formattedText, 0, 0);
            workbook.CurrentWorksheet.SetRowHeight(0, 40);
            workbook.CurrentWorksheet.AddCell("Raben", 0, 1);
            workbook.SaveAs(@"C:\purge-temp\formattedTextTest.xlsx");
            Workbook wb2 = WorkbookReader.Load(@"C:\purge-temp\formattedTextTest.xlsx");
            // wb2.CurrentWorksheet.Cells["A1"]
            int i = 0;
        }
    }
}
