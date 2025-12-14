using NanoXLSX.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX.Formatting.Test.Worksheets
{
    public class WorksheetTest
    {
        [Fact]
        public void SampleTestMethod()
        {
            Workbook workbook = new Workbook("worksheet1");
            FormattedText formattedText = new FormattedText();
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyle style1 = builder.Font("Arial").Bold().ColorRgb("CC00FF").Build();
            InlineStyle style2 = builder.Font("Calibri").Italic().ColorRgb("00FFFF").Build();
            formattedText.AddRun("Line1", style1);
            formattedText.AddRun(" Line2", style2);
            workbook.CurrentWorksheet.AddCell(formattedText, 0, 0);
        }
    }
}
