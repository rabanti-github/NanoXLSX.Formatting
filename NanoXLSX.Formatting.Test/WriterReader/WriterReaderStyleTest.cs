using NanoXLSX.Extensions;
using NanoXLSX.Formatting.Test;
using NanoXLSX.Registry;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.Diagnostics.SymbolStore;
using System.IO;
using Xunit;
using static System.Net.Mime.MediaTypeNames;

namespace NanoXLSX.Tests
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
    public class WriterReaderStyleTest : IDisposable
    {

       public WriterReaderStyleTest()
        {
            TestUtils.InitializePlugIns();
        }

        #region FormattedText with Font Styles

        [Fact(DisplayName = "Test of writing and reading FormattedText with bold font")]
        public void WriteReadFormattedTextBoldFontTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            Font boldFont = new Font();
            boldFont.Bold = true;
            originalText.AddRun("Bold text", boldFont);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);
            workbook.SaveAs(@"C:\purge-temp\NanoXLSX.Formatting\bold.xlsx");
            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.NotNull(loadedText);
            Assert.Equal("Bold text", loadedText.PlainText);
            Assert.NotNull(loadedText.Runs[0].FontStyle);
            Assert.True(loadedText.Runs[0].FontStyle.Bold);
            Assert.Equal(boldFont.GetHashCode(), loadedText.Runs[0].FontStyle.GetHashCode());
        }

        [Fact(DisplayName = "Test of writing and reading FormattedText with italic font")]
        public void WriteReadFormattedTextItalicFontTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            Font italicFont = new Font();
            italicFont.Italic = true;
            originalText.AddRun("Italic text", italicFont);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.NotNull(loadedText);
            Assert.NotNull(loadedText.Runs[0].FontStyle);
            Assert.True(loadedText.Runs[0].FontStyle.Italic);
            Assert.Equal(italicFont.GetHashCode(), loadedText.Runs[0].FontStyle.GetHashCode());
        }

        [Fact(DisplayName = "Test of writing and reading FormattedText with underline font")]
        public void WriteReadFormattedTextUnderlineFontTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            Font underlineFont = new Font();
            underlineFont.Underline = Font.UnderlineValue.Single;
            originalText.AddRun("Underlined text", underlineFont);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.NotNull(loadedText);
            Assert.NotNull(loadedText.Runs[0].FontStyle);
            Assert.Equal(Font.UnderlineValue.Single, loadedText.Runs[0].FontStyle.Underline);
            Assert.Equal(underlineFont.GetHashCode(), loadedText.Runs[0].FontStyle.GetHashCode());
        }

        [Theory(DisplayName = "Test of writing and reading FormattedText with various underline styles")]
        [InlineData(Font.UnderlineValue.Single)]
        [InlineData(Font.UnderlineValue.Double)]
        [InlineData(Font.UnderlineValue.SingleAccounting)]
        [InlineData(Font.UnderlineValue.DoubleAccounting)]
        public void WriteReadFormattedTextVariousUnderlineStylesTest(Font.UnderlineValue underlineStyle)
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            Font font = new Font();
            font.Underline = underlineStyle;
            originalText.AddRun("Underlined", font);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.Equal(underlineStyle, loadedText.Runs[0].FontStyle.Underline);
        }

        [Fact(DisplayName = "Test of writing and reading FormattedText with strikethrough font")]
        public void WriteReadFormattedTextStrikethroughFontTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            Font strikeFont = new Font();
            strikeFont.Strike = true;
            originalText.AddRun("Strikethrough text", strikeFont);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.NotNull(loadedText);
            Assert.NotNull(loadedText.Runs[0].FontStyle);
            Assert.True(loadedText.Runs[0].FontStyle.Strike);
            Assert.Equal(strikeFont.GetHashCode(), loadedText.Runs[0].FontStyle.GetHashCode());
        }

        [Theory(DisplayName = "Test of writing and reading FormattedText with various font sizes")]
        [InlineData(8)]
        [InlineData(10)]
        [InlineData(12)]
        [InlineData(14)]
        [InlineData(18)]
        [InlineData(24)]
        public void WriteReadFormattedTextVariousFontSizesTest(float size)
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            Font font = new Font();
            font.Size = size;
            originalText.AddRun("Sized text", font);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.NotNull(loadedText);
            Assert.NotNull(loadedText.Runs[0].FontStyle);
            Assert.Equal(size, loadedText.Runs[0].FontStyle.Size);
            Assert.Equal(font.GetHashCode(), loadedText.Runs[0].FontStyle.GetHashCode());
        }

        [Theory(DisplayName = "Test of writing and reading FormattedText with various font names")]
        [InlineData("Arial")]
        [InlineData("Calibri")]
        [InlineData("Times New Roman")]
        [InlineData("Courier New")]
        public void WriteReadFormattedTextVariousFontNamesTest(string fontName)
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            Font font = new Font();
            font.Name = fontName;
            originalText.AddRun("Font name test", font);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.NotNull(loadedText);
            Assert.NotNull(loadedText.Runs[0].FontStyle);
            Assert.Equal(fontName, loadedText.Runs[0].FontStyle.Name);
            Assert.Equal(font.GetHashCode(), loadedText.Runs[0].FontStyle.GetHashCode());
        }

        [Fact(DisplayName = "Test of writing and reading FormattedText with multiple font properties")]
        public void WriteReadFormattedTextMultipleFontPropertiesTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            Font font = new Font();
            font.Bold = true;
            font.Italic = true;
            font.Size = 16;
            font.Name = "Calibri";
            font.Underline = Font.UnderlineValue.Double;
            font.Strike = true;
            originalText.AddRun("Complex styled text", font);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.NotNull(loadedText);
            Font loadedFont = loadedText.Runs[0].FontStyle;
            Assert.NotNull(loadedFont);
            Assert.True(loadedFont.Bold);
            Assert.True(loadedFont.Italic);
            Assert.Equal(16, loadedFont.Size);
            Assert.Equal("Calibri", loadedFont.Name);
            Assert.Equal(Font.UnderlineValue.Double, loadedFont.Underline);
            Assert.True(loadedFont.Strike);
            Assert.Equal(font.GetHashCode(), loadedFont.GetHashCode());
        }

        [Fact(DisplayName = "Test of writing and reading FormattedText with mixed styled runs")]
        public void WriteReadFormattedTextMixedStyledRunsTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            Font boldFont = new Font();
            boldFont.Bold = true;
            Font italicFont = new Font();
            italicFont.Italic = true;

            originalText.AddRun("Bold", boldFont);
            originalText.AddRun(" ");
            originalText.AddRun("Italic", italicFont);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.NotNull(loadedText);
            Assert.Equal("Bold Italic", loadedText.PlainText);
            Assert.Equal(3, loadedText.Runs.Count);
            Assert.True(loadedText.Runs[0].FontStyle.Bold);
            Assert.Null(loadedText.Runs[1].FontStyle);
            Assert.True(loadedText.Runs[2].FontStyle.Italic);
        }

        [Fact(DisplayName = "Test of writing and reading FormattedText with run without font style")]
        public void WriteReadFormattedTextRunWithoutFontStyleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            originalText.AddRun("Plain text");
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.NotNull(loadedText);
            Assert.Equal("Plain text", loadedText.PlainText);
            Assert.Null(loadedText.Runs[0].FontStyle);
        }

        [Fact(DisplayName = "Test of writing and reading FormattedText with superscript")]
        public void WriteReadFormattedTextSuperscriptTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            Font font = new Font();
            font.VerticalAlign = Font.VerticalTextAlignValue.Superscript;
            originalText.AddRun("x²", font);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.NotNull(loadedText);
            Assert.Equal(Font.VerticalTextAlignValue.Superscript, loadedText.Runs[0].FontStyle.VerticalAlign);
            Assert.Equal(font.GetHashCode(), loadedText.Runs[0].FontStyle.GetHashCode());
        }

        [Fact(DisplayName = "Test of writing and reading FormattedText with subscript")]
        public void WriteReadFormattedTextSubscriptTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            Font font = new Font();
            font.VerticalAlign = Font.VerticalTextAlignValue.Subscript;
            originalText.AddRun("H₂O", font);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.NotNull(loadedText);
            Assert.Equal(Font.VerticalTextAlignValue.Subscript, loadedText.Runs[0].FontStyle.VerticalAlign);
            Assert.Equal(font.GetHashCode(), loadedText.Runs[0].FontStyle.GetHashCode());
        }

        #endregion

        #region FormattedText with Line Breaks

        [Fact(DisplayName = "Test of writing and reading FormattedText with line break")]
        public void WriteReadFormattedTextWithLineBreakTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            originalText.AddRun("Line 1");
            originalText.AddLineBreak();
            originalText.AddRun("Line 2");
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.NotNull(loadedText);
            Assert.Contains("\n", loadedText.PlainText);
            Assert.True(loadedText.WrapText);
        }

        [Fact(DisplayName = "Test of writing and reading FormattedText with multiple line breaks")]
        public void WriteReadFormattedTextWithMultipleLineBreaksTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            originalText.AddRun("Line 1");
            originalText.AddLineBreak();
            originalText.AddRun("Line 2");
            originalText.AddLineBreak();
            originalText.AddRun("Line 3");
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.NotNull(loadedText);
            Assert.True(loadedText.WrapText);
            string[] lines = loadedText.PlainText.Split(new[] { "\n" }, StringSplitOptions.None);
            Assert.Equal(3, lines.Length);
            Assert.Equal("Line 1", lines[0]);
            Assert.Equal("Line 2", lines[1]);
            Assert.Equal("Line 3", lines[2]);
        }

        [Fact(DisplayName = "Test of writing and reading FormattedText with WrapText applies cell style")]
        public void WriteReadFormattedTextWrapTextAppliesCellStyleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            originalText.AddRun("Line 1");
            originalText.AddLineBreak();
            originalText.AddRun("Line 2");
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            Assert.NotNull(cell.CellStyle);
            Assert.Equal(CellXf.TextBreakValue.WrapText, cell.CellStyle.CurrentCellXf.Alignment);
        }

        [Fact(DisplayName = "Test of writing and reading FormattedText with line break and custom style")]
        public void WriteReadFormattedTextLineBreakWithCustomStyleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            originalText.AddRun("Line 1");
            originalText.AddLineBreak();
            originalText.AddRun("Line 2");
            Style customStyle = new Style();
            customStyle.CurrentFont.Bold = true;
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0, customStyle);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            Assert.NotNull(cell.CellStyle);
            Assert.Equal(CellXf.TextBreakValue.WrapText, cell.CellStyle.CurrentCellXf.Alignment);
            Assert.True(cell.CellStyle.CurrentFont.Bold);
        }

        [Fact(DisplayName = "Test of writing and reading FormattedText with line break in middle of text")]
        public void WriteReadFormattedTextLineBreakInMiddleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            originalText.AddRun("Start");
            originalText.AddLineBreak();
            originalText.AddRun("Middle");
            originalText.AddLineBreak();
            originalText.AddRun("End");
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.Contains("Start", loadedText.PlainText);
            Assert.Contains("Middle", loadedText.PlainText);
            Assert.Contains("End", loadedText.PlainText);
        }

        #endregion

        #region FormattedText at Different Positions

        [Theory(DisplayName = "Test of writing and reading FormattedText at various positions")]
        [InlineData(0, 0)]
        [InlineData(5, 10)]
        [InlineData(100, 200)]
        public void WriteReadFormattedTextAtVariousPositionsTest(int column, int row)
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            originalText.AddRun($"Text at {column},{row}");
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, column, row);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(column, row);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.NotNull(loadedText);
            Assert.Equal($"Text at {column},{row}", loadedText.PlainText);
        }

        [Fact(DisplayName = "Test of writing and reading multiple FormattedText cells")]
        public void WriteReadMultipleFormattedTextCellsTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText text1 = new FormattedText().AddRun("First");
            FormattedText text2 = new FormattedText().AddRun("Second");
            FormattedText text3 = new FormattedText().AddRun("Third");

            workbook.CurrentWorksheet.AddFormattedTextCell(text1, 0, 0);
            workbook.CurrentWorksheet.AddFormattedTextCell(text2, 1, 0);
            workbook.CurrentWorksheet.AddFormattedTextCell(text3, 2, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            FormattedText loaded1 = loadedWorkbook.CurrentWorksheet.GetCell(0, 0).Value as FormattedText;
            FormattedText loaded2 = loadedWorkbook.CurrentWorksheet.GetCell(1, 0).Value as FormattedText;
            FormattedText loaded3 = loadedWorkbook.CurrentWorksheet.GetCell(2, 0).Value as FormattedText;

            Assert.Equal("First", loaded1.PlainText);
            Assert.Equal("Second", loaded2.PlainText);
            Assert.Equal("Third", loaded3.PlainText);
        }

        [Fact(DisplayName = "Test of writing and reading FormattedText cells in grid pattern")]
        public void WriteReadFormattedTextCellsGridPatternTest()
        {
            Workbook workbook = new Workbook("sheet1");

            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    FormattedText text = new FormattedText();
                    text.AddRun($"R{row}C{col}");
                    workbook.CurrentWorksheet.AddFormattedTextCell(text, col, row);
                }
            }

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(col, row);
                    FormattedText loadedText = cell.Value as FormattedText;
                    Assert.NotNull(loadedText);
                    Assert.Equal($"R{row}C{col}", loadedText.PlainText);
                }
            }
        }

        #endregion

        #region FormattedText with PhoneticRuns

        [Fact(DisplayName = "Test of writing and reading FormattedText with single phonetic run")]
        public void WriteReadFormattedTextWithSinglePhoneticRunTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            originalText.AddRun("漢字");
            originalText.AddPhoneticRun("かんじ", 0, 2);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.NotNull(loadedText);
            Assert.Single(loadedText.PhoneticRuns);
            Assert.Equal("かんじ", loadedText.PhoneticRuns[0].Text);
            Assert.Equal(0u, loadedText.PhoneticRuns[0].StartBase);
            Assert.Equal(2u, loadedText.PhoneticRuns[0].EndBase);
        }

        [Fact(DisplayName = "Test of writing and reading FormattedText with multiple phonetic runs")]
        public void WriteReadFormattedTextWithMultiplePhoneticRunsTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            originalText.AddRun("東京都");
            originalText.AddPhoneticRun("とう", 0, 1);
            originalText.AddPhoneticRun("きょう", 1, 2);
            originalText.AddPhoneticRun("と", 2, 3);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.NotNull(loadedText);
            Assert.Equal(3, loadedText.PhoneticRuns.Count);
            Assert.Equal("とう", loadedText.PhoneticRuns[0].Text);
            Assert.Equal("きょう", loadedText.PhoneticRuns[1].Text);
            Assert.Equal("と", loadedText.PhoneticRuns[2].Text);
        }

        [Fact(DisplayName = "Test of writing and reading FormattedText with phonetic properties")]
        public void WriteReadFormattedTextWithPhoneticPropertiesTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            Font phoneticFont = new Font();
            phoneticFont.Size = 8;
            originalText.AddRun("漢字");
            originalText.AddPhoneticRun("かんじ", 0, 2);
            originalText.SetPhoneticProperties(phoneticFont, PhoneticRun.PhoneticType.Hiragana, PhoneticRun.PhoneticAlignment.Center);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);
            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.NotNull(loadedText);
            Assert.NotNull(loadedText.PhoneticProperties);
            Assert.Equal(PhoneticRun.PhoneticType.Hiragana, loadedText.PhoneticProperties.Type);
            Assert.Equal(PhoneticRun.PhoneticAlignment.Center, loadedText.PhoneticProperties.Alignment);
        }

        [Theory(DisplayName = "Test of writing and reading FormattedText with various phonetic types")]
        [InlineData(PhoneticRun.PhoneticType.FullwidthKatakana)]
        [InlineData(PhoneticRun.PhoneticType.HalfwidthKatakana)]
        [InlineData(PhoneticRun.PhoneticType.Hiragana)]
        [InlineData(PhoneticRun.PhoneticType.NoConversion)]
        public void WriteReadFormattedTextVariousPhoneticTypesTest(PhoneticRun.PhoneticType type)
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            Font phoneticFont = new Font();
            originalText.AddRun("日本");
            originalText.AddPhoneticRun("にほん", 0, 2);
            originalText.SetPhoneticProperties(phoneticFont, type);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.Equal(type, loadedText.PhoneticProperties.Type);
        }

        [Theory(DisplayName = "Test of writing and reading FormattedText with various phonetic alignments")]
        [InlineData(PhoneticRun.PhoneticAlignment.NoControl)]
        [InlineData(PhoneticRun.PhoneticAlignment.Left)]
        [InlineData(PhoneticRun.PhoneticAlignment.Center)]
        [InlineData(PhoneticRun.PhoneticAlignment.Distributed)]
        public void WriteReadFormattedTextVariousPhoneticAlignmentsTest(PhoneticRun.PhoneticAlignment alignment)
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            Font phoneticFont = new Font();
            originalText.AddRun("日本");
            originalText.AddPhoneticRun("にほん", 0, 2);
            originalText.SetPhoneticProperties(phoneticFont, PhoneticRun.PhoneticType.Hiragana, alignment);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.Equal(alignment, loadedText.PhoneticProperties.Alignment);
        }

        #endregion

        #region FormattedText Complex Scenarios

        [Fact(DisplayName = "Test of writing and reading FormattedText with all features combined")]
        public void WriteReadFormattedTextAllFeaturesCombinedTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();

            Font boldFont = new Font();
            boldFont.Bold = true;
            boldFont.Size = 14;

            Font italicFont = new Font();
            italicFont.Italic = true;
            italicFont.Name = "Arial";

            Font phoneticFont = new Font();
            phoneticFont.Size = 8;

            originalText.AddRun("Bold", boldFont);
            originalText.AddRun(" ");
            originalText.AddRun("Italic", italicFont);
            originalText.AddLineBreak();
            originalText.AddRun("漢字", boldFont);
            originalText.AddPhoneticRun("かんじ", 0, 2);
            originalText.SetPhoneticProperties(phoneticFont, PhoneticRun.PhoneticType.Hiragana, PhoneticRun.PhoneticAlignment.Center);

            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.NotNull(loadedText);
            Assert.True(loadedText.WrapText);
            Assert.Contains("\n", loadedText.PlainText);
            Assert.Equal(4, loadedText.Runs.Count);
            Assert.True(loadedText.Runs[0].FontStyle.Bold);
            Assert.True(loadedText.Runs[2].FontStyle.Italic);
            Assert.Single(loadedText.PhoneticRuns);
            Assert.NotNull(loadedText.PhoneticProperties);
        }

        [Fact(DisplayName = "Test of writing and reading FormattedText with long text content")]
        public void WriteReadFormattedTextLongTextTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();
            string longText = new string('A', 1000);
            originalText.AddRun(longText);
            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.NotNull(loadedText);
            Assert.Equal(1000, loadedText.PlainText.Length);
            Assert.Equal(longText, loadedText.PlainText);
        }

        [Fact(DisplayName = "Test of writing and reading FormattedText with many runs")]
        public void WriteReadFormattedTextManyRunsTest()
        {
            Workbook workbook = new Workbook("sheet1");
            FormattedText originalText = new FormattedText();

            for (int i = 0; i < 20; i++)
            {
                Font font = new Font();
                font.Bold = i % 2 == 0;
                originalText.AddRun($"Run{i}", font);
            }

            workbook.CurrentWorksheet.AddFormattedTextCell(originalText, 0, 0);

            Workbook loadedWorkbook = SaveAndReadWorkbook(workbook);

            Cell cell = loadedWorkbook.CurrentWorksheet.GetCell(0, 0);
            FormattedText loadedText = cell.Value as FormattedText;
            Assert.NotNull(loadedText);
            Assert.Equal(20, loadedText.Runs.Count);

            for (int i = 0; i < 20; i++)
            {
                Assert.Equal($"Run{i}", loadedText.Runs[i].Text);
                Assert.Equal(i % 2 == 0, loadedText.Runs[i].FontStyle.Bold);
            }
        }

        #endregion

        private Workbook SaveAndReadWorkbook(Workbook workbook)
        {
            MemoryStream stream = new MemoryStream();
            workbook.SaveAsStream(stream, true);
            stream.Position = 0;
            Workbook loadedWorkbook = WorkbookReader.Load(stream);
            return loadedWorkbook;
        }

        private FormattedText GetFormattedText(List<Tuple<string, Font>> runs)
        {
            FormattedText text = new FormattedText();
            foreach(Tuple<string, Font> run in runs)
            {
                if (run.Item2 == null)
                {
                    text.AddRun(run.Item1);
                }
                else
                {
                    text.AddRun(run.Item1, run.Item2);
                }
            }
            return text;
        }

        public void Dispose()
        {
            TestUtils.DisposePlugIns();
        }
    }
}