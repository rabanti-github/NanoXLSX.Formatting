using NanoXLSX.Styles;
using System;
using Xunit;

namespace NanoXLSX.Formatting.Test.WriterReader
{
    /// <summary>
    /// Tests where a written formatted text do not carry formatting information. They act like simple strings, but may be defined as runs.
    /// </summary>
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
    public class WriterReaderSimpleTest : IDisposable
    {
        public WriterReaderSimpleTest()
        {
            TestUtils.InitializePlugIns();
        }

        [Theory(DisplayName = "Test of writing and reading FormattedText with single run, without style")]
        [InlineData("Simple text")]
        [InlineData("Text with numbers 123")]
        [InlineData("Special chars: @#$%^&*()")]
        [InlineData("Unicode: 日本語")]
        [InlineData("Emoji: 😀🎉")]
        [InlineData(" ")]
        [InlineData("   ")]
        [InlineData("\t\t")]
        public void WriteReadFormattedTextSingleRunNoStyleTest(string text)
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText().AddRun(text);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);
            Workbook loadedWorkbook = TestUtils.SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            Assert.NotNull(cell);
            Assert.Equal(Cell.CellType.String, cell.DataType);
            Assert.IsType<FormattedText>(cell.Value);
            FormattedText expectedText = cell.Value as FormattedText;
            Assert.Equal(text, expectedText.PlainText);
            Assert.Null(cell.CellStyle);
        }

        [Theory(DisplayName = "Test of writing and reading FormattedText with single run and normal text, without style")]
        [InlineData("Simple text")]
        [InlineData("Text with numbers 123")]
        [InlineData("Special chars: @#$%^&*()")]
        [InlineData("Unicode: 日本語")]
        [InlineData("Emoji: 😀🎉")]
        [InlineData(" ")]
        [InlineData("   ")]
        [InlineData("\t\t")]
        public void WriteReadFormattedTextSingleRunAndTextNoStyleTest(string text)
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText().AddRun(text);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);
            workbook.CurrentWorksheet.AddCell(text, 0, 1);
            Workbook loadedWorkbook = TestUtils.SaveAndReadWorkbook(workbook);

            Cell cell1 = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            Cell cell2 = loadedWorkbook.CurrentWorksheet.GetCell(0, 1);
            Assert.NotNull(cell1);
            Assert.NotNull(cell2);
            Assert.Equal(Cell.CellType.String, cell1.DataType);
            Assert.Equal(Cell.CellType.String, cell2.DataType);
            Assert.IsType<FormattedText>(cell1.Value);
            Assert.IsType<string>(cell2.Value);
            FormattedText expectedText1 = cell1.Value as FormattedText;
            string expectedText2 = cell2.Value as string;
            Assert.Equal(text, expectedText1.PlainText);
            Assert.Equal(text, expectedText2);
            Assert.Null(cell1.CellStyle);
            Assert.Null(cell2.CellStyle);
        }


        [Theory(DisplayName = "Test of writing and reading FormattedText with multiple runs, without style")]
        [InlineData("Simple ", "text")]
        [InlineData("Text with numbers ", "123")]
        [InlineData("Special chars: ", "@#$%^&*()")]
        [InlineData(" ", "日本語")]
        [InlineData(" ", "\t")]
        [InlineData("\t", "\t")]
        public void WriteReadFormattedTextMultipleRunsNoStyleTest(string text1, string text2)
        {
            string expectedText = text1 + text2;
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            originalText.AddRun(text1);
            originalText.AddRun(text2);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);
            Workbook loadedWorkbook = TestUtils.SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            Assert.Equal(Cell.CellType.String, cell.DataType);
            Assert.IsType<FormattedText>(cell.Value);
            Assert.Equal(expectedText, (cell.Value as FormattedText).PlainText);
            Assert.Null(cell.CellStyle);
        }

        [Theory(DisplayName = "Test of writing and reading FormattedText with one run, whether text wrapping is identified, with no additional style")]
        [InlineData("Simple Text", false, "Simple Text")]
        [InlineData("Simple\nText", true, "Simple\nText")]
        [InlineData("Simple\r\nText", true, "Simple\nText")]
        [InlineData("  ", false, "  ")]
        [InlineData("\n\n", true, "\n\n")]
        [InlineData("\ntext\n", true, "\ntext\n")]
        [InlineData("\r\n\r\n", true, "\n\n")]
        [InlineData("\r\n0", true, "\n0")]
        public void WriteReadFormattedTextHasWrapStyleTest(string givenText, bool expectWrapStyle, string expectedText)
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            originalText.AddRun(givenText);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);
            Workbook loadedWorkbook = TestUtils.SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            Assert.Equal(Cell.CellType.String, cell.DataType);
            Assert.IsType<FormattedText>(cell.Value);
            Assert.Equal(expectedText, (cell.Value as FormattedText).PlainText);
            if (expectWrapStyle)
            {
                Assert.NotNull(cell.CellStyle);
                Assert.Equal(CellXf.TextBreakValue.WrapText, cell.CellStyle.CurrentCellXf.Alignment);
            }
            else
            {
                Assert.Null(cell.CellStyle);
            }
        }

        [Theory(DisplayName = "Test of writing and reading FormattedText with multiple runs whether text wrapping is identified, with no additional style")]
        [InlineData("Simple ", "text", false, "Simple text")]
        [InlineData("Text with numbers\n", "123", true, "Text with numbers\n123")]
        [InlineData("Special chars:\r\n", "@#$%^&*()", true, "Special chars:\n@#$%^&*()")]
        [InlineData(" ", "日本語", false, " 日本語")]
        [InlineData("\n\r\n", "😀🎉", true, "\n\n😀🎉")]
        [InlineData(" ", "\t", false, " \t")]
        [InlineData("   ", "\n", true, "   \n")]
        [InlineData("\t", "\t", false, "\t\t")]
        [InlineData("\n\r", "\r\n", true, "\n\n")]
        [InlineData("\n\r\rtext", " ", true, "\n\ntext ")]// \r\n is one instance, remaining \r is second
        public void WriteReadFormattedTextMultipleRunsHasWrapStyleTest(string givenText1, string givenText2, bool expectWrapStyle, string expectedText)
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            originalText.AddRun(givenText1);
            originalText.AddRun(givenText2);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);
            Workbook loadedWorkbook = TestUtils.SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            Assert.Equal(Cell.CellType.String, cell.DataType);
            Assert.IsType<FormattedText>(cell.Value);
            Assert.Equal(expectedText, (cell.Value as FormattedText).PlainText);
            if (expectWrapStyle)
            {
                Assert.NotNull(cell.CellStyle);
                Assert.Equal(CellXf.TextBreakValue.WrapText, cell.CellStyle.CurrentCellXf.Alignment);
            }
            else
            {
                Assert.Null(cell.CellStyle);
            }
        }

        [Theory(DisplayName = "Test of writing and reading FormattedText with one run, whether text wrapping is identified, with additional style")]
        [InlineData("Simple Text", false, "Simple Text")]
        [InlineData("Simple\nText", true, "Simple\nText")]
        [InlineData("Simple\r\nText", true, "Simple\nText")]
        [InlineData("  ", false, "  ")]
        [InlineData("\n\n", true, "\n\n")]
        [InlineData("\r\n\r\n", true, "\n\n")]
        [InlineData("\r\n0", true, "\n0")]
        public void WriteReadFormattedTextHasWrapMixedStyleTest(string givenText, bool expectWrapStyle, string expectedText)
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            originalText.AddRun(givenText);
            Style style = BasicStyles.Bold;
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0, style);
            Workbook loadedWorkbook = TestUtils.SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            Assert.Equal(Cell.CellType.String, cell.DataType);
            Assert.IsType<FormattedText>(cell.Value);
            Assert.Equal(expectedText, (cell.Value as FormattedText).PlainText);
            Assert.NotNull(cell.CellStyle);
            Assert.True(cell.CellStyle.CurrentFont.Bold);
            if (expectWrapStyle)
            {
                Assert.Equal(CellXf.TextBreakValue.WrapText, cell.CellStyle.CurrentCellXf.Alignment);
            }
            else
            {
                Assert.Equal(CellXf.TextBreakValue.None, cell.CellStyle.CurrentCellXf.Alignment);
            }
        }

        [Fact(DisplayName = "Test of an empty formatted text object (should create an empty string)")]
        public void EmptyRunTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);
            Workbook loadedWorkbook = TestUtils.SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            Assert.Equal(Cell.CellType.String, cell.DataType);
            Assert.IsType<string>(cell.Value);
            Assert.Equal("", cell.Value as string);
            Assert.Null(cell.CellStyle);
        }

        [Fact(DisplayName = "Test of an empty formatted text object but with phonetic info")]
        public void EmptyRunWithWithPhoneticInfoTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            originalText.SetPhoneticProperties(null, Extensions.PhoneticRun.PhoneticType.Hiragana, Extensions.PhoneticRun.PhoneticAlignment.Distributed);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);
            Workbook loadedWorkbook = TestUtils.SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            Assert.Equal(Cell.CellType.String, cell.DataType);
            Assert.IsType<string>(cell.Value);
            Assert.Equal("", cell.Value as string);
            Assert.Null(cell.CellStyle);
        }

        public void Dispose()
        {
            TestUtils.DisposePlugIns();
        }
    }
}