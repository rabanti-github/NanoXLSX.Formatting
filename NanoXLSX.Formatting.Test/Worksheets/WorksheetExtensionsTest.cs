using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using Xunit;

namespace NanoXLSX.Formatting.Test
{
    public class WorksheetExtensionsTest
    {
        [Fact(DisplayName = "Test of AddFormattedTextCell with column and row numbers")]
        public void AddFormattedTextCellWithColumnAndRowTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText formattedText = new FormattedText();
            formattedText.AddRun("Test");
            worksheet.AddFormattedTextCell(formattedText, 0, 0);
            Cell cell = worksheet.GetCell(0, 0);
            Assert.NotNull(cell);
            Assert.IsType<FormattedText>(cell.Value);
        }

        [Fact(DisplayName = "Test of AddFormattedTextCell with column, row and style")]
        public void AddFormattedTextCellWithColumnRowAndStyleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText formattedText = new FormattedText();
            formattedText.AddRun("Test");
            Style style = new Style();
            style.CurrentFont.Bold = true;
            worksheet.AddFormattedTextCell(formattedText, 0, 0, style);
            Cell cell = worksheet.GetCell(0, 0);
            Assert.NotNull(cell);
            Assert.NotNull(cell.CellStyle);
            Assert.Equal(style.GetHashCode(), cell.CellStyle.GetHashCode());
        }

        [Fact(DisplayName = "Test of AddFormattedTextCell with WrapText true applies LineBreakStyle")]
        public void AddFormattedTextCellWithWrapTextAppliesLineBreakStyleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText formattedText = new FormattedText();
            formattedText.AddRun("Line1");
            formattedText.AddLineBreak();
            formattedText.AddRun("Line2");
            worksheet.AddFormattedTextCell(formattedText, 0, 0);
            Cell cell = worksheet.GetCell(0, 0);
            Assert.NotNull(cell.CellStyle);
            Assert.Equal(CellXf.TextBreakValue.WrapText, cell.CellStyle.CurrentCellXf.Alignment);
        }

        [Fact(DisplayName = "Test of AddFormattedTextCell with WrapText true merges styles")]
        public void AddFormattedTextCellWithWrapTextMergesStylesTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText formattedText = new FormattedText();
            formattedText.AddRun("Line1");
            formattedText.AddLineBreak();
            Style style = new Style();
            style.CurrentFont.Bold = true;
            worksheet.AddFormattedTextCell(formattedText, 0, 0, style);
            Cell cell = worksheet.GetCell(0, 0);
            Assert.NotNull(cell.CellStyle);
            Assert.True(cell.CellStyle.CurrentFont.Bold);
            Assert.Equal(CellXf.TextBreakValue.WrapText, cell.CellStyle.CurrentCellXf.Alignment);
        }

        [Fact(DisplayName = "Test of the failing AddFormattedTextCell when no worksheet was defined - exception NullReferenceException")]
        public void AddFormattedTextCellWithNullWorksheetFailingTest()
        {
            Workbook workbook = new Workbook(); // Default constructor does not create a worksheet
            Assert.Throws<NullReferenceException>(() => workbook.CurrentWorksheet.AddFormattedTextCell(new FormattedText(), 0, 0));
        }

        [Fact(DisplayName = "Test of AddFormattedTextCell with null FormattedText throws WorksheetException")]
        public void AddFormattedTextCellWithNullFormattedTextFailingTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            Assert.Throws<WorksheetException>(() => worksheet.AddFormattedTextCell(null, 0, 0));
        }

        [Fact(DisplayName = "Test of AddFormattedTextCell with null FormattedText and style throws WorksheetException")]
        public void AddFormattedTextCellWithNullFormattedTextAndStyleFailingTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            Style style = new Style();
            Assert.Throws<WorksheetException>(() => worksheet.AddFormattedTextCell(null, 0, 0, style));
        }

        [Theory(DisplayName = "Test of AddFormattedTextCell with various column and row positions")]
        [InlineData(0, 0)]
        [InlineData(5, 10)]
        [InlineData(100, 200)]
        public void AddFormattedTextCellWithVariousPositionsTest(int column, int row)
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText formattedText = new FormattedText();
            formattedText.AddRun("Test");
            worksheet.AddFormattedTextCell(formattedText, column, row);
            Cell cell = worksheet.GetCell(column, row);
            Assert.NotNull(cell);
        }

        [Theory(DisplayName = "Test of AddFormattedTextCell with address string")]
        [InlineData("A1")]
        [InlineData("B5")]
        [InlineData("Z10")]
        [InlineData("a1")]
        [InlineData("$B$5")]
        [InlineData("AA$5")]
        [InlineData("$F99800")]
        [InlineData("XFD1048576")]
        public void AddFormattedTextCellWithAddressStringTest(string address)
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText formattedText = new FormattedText();
            formattedText.AddRun("Test");
            worksheet.AddFormattedTextCell(formattedText, address);
            int column, row;
            Cell.ResolveCellCoordinate(address, out column, out row);
            Cell cell = worksheet.GetCell(column, row);
            Assert.NotNull(cell);
        }

        [Fact(DisplayName = "Test of AddFormattedTextCell with address string and style")]
        public void AddFormattedTextCellWithAddressStringAndStyleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText formattedText = new FormattedText();
            formattedText.AddRun("Test");
            Style style = new Style();
            style.CurrentFont.Italic = true;
            worksheet.AddFormattedTextCell(formattedText, "A1", style);
            Cell cell = worksheet.GetCell(0, 0);
            Assert.NotNull(cell);
            Assert.NotNull(cell.CellStyle);
            Assert.True(cell.CellStyle.CurrentFont.Italic);
        }

        [Fact(DisplayName = "Test of AddFormattedTextCell with null FormattedText and address throws WorksheetException")]
        public void AddFormattedTextCellWithNullFormattedTextAndAddressFailingTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            Assert.Throws<WorksheetException>(() => worksheet.AddFormattedTextCell(null, "A1"));
        }

        [Fact(DisplayName = "Test of AddFormattedTextCell with null FormattedText, address and style throws WorksheetException")]
        public void AddFormattedTextCellWithNullFormattedTextAddressAndStyleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            Style style = new Style();
            Assert.Throws<WorksheetException>(() => worksheet.AddFormattedTextCell(null, "A1", style));
        }

        [Fact(DisplayName = "Test of AddNextFormattedTextCell without style")]
        public void AddNextFormattedTextCellWithoutStyleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText formattedText = new FormattedText();
            formattedText.AddRun("Test");
            worksheet.AddNextFormattedTextCell(formattedText);
            Cell cell = worksheet.GetCell(0, 0);
            Assert.NotNull(cell);
        }

        [Fact(DisplayName = "Test of AddNextFormattedTextCell with style")]
        public void AddNextFormattedTextCellWithStyleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText formattedText = new FormattedText();
            formattedText.AddRun("Test");
            Style style = new Style();
            style.CurrentFont.Bold = true;
            worksheet.AddNextFormattedTextCell(formattedText, style);
            Cell cell = worksheet.GetCell(0, 0);
            Assert.NotNull(cell);
            Assert.NotNull(cell.CellStyle);
        }

        [Fact(DisplayName = "Test of AddNextFormattedTextCell with WrapText true applies LineBreakStyle")]
        public void AddNextFormattedTextCellWithWrapTextAppliesLineBreakStyleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText formattedText = new FormattedText();
            formattedText.AddRun("Line1");
            formattedText.AddLineBreak();
            worksheet.AddNextFormattedTextCell(formattedText);
            Cell cell = worksheet.GetCell(0, 0);
            Assert.NotNull(cell.CellStyle);
            Assert.Equal(CellXf.TextBreakValue.WrapText, cell.CellStyle.CurrentCellXf.Alignment);
        }

        [Fact(DisplayName = "Test of AddNextFormattedTextCell with WrapText true merges styles")]
        public void AddNextFormattedTextCellWithWrapTextMergesStylesTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText formattedText = new FormattedText();
            formattedText.AddRun("Line1");
            formattedText.AddLineBreak();
            Style style = new Style();
            style.CurrentFont.Italic = true;
            worksheet.AddNextFormattedTextCell(formattedText, style);
            Cell cell = worksheet.GetCell(0, 0);
            Assert.NotNull(cell.CellStyle);
            Assert.True(cell.CellStyle.CurrentFont.Italic);
            Assert.Equal(CellXf.TextBreakValue.WrapText, cell.CellStyle.CurrentCellXf.Alignment);
        }

        [Fact(DisplayName = "Test of AddNextFormattedTextCell with null FormattedText throws WorksheetException")]
        public void AddNextFormattedTextCellWithNullFormattedTextFailingTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            Assert.Throws<WorksheetException>(() => worksheet.AddNextFormattedTextCell(null));
        }

        [Fact(DisplayName = "Test of AddNextFormattedTextCell with null FormattedText and style throws WorksheetException")]
        public void AddNextFormattedTextCellWithNullFormattedTextAndStyleFailingTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            Style style = new Style();
            Assert.Throws<WorksheetException>(() => worksheet.AddNextFormattedTextCell(null, style));
        }

        [Fact(DisplayName = "Test of AddNextFormattedTextCell increments position")]
        public void AddNextFormattedTextCellIncrementsPositionTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText text1 = new FormattedText();
            text1.AddRun("First");
            FormattedText text2 = new FormattedText();
            text2.AddRun("Second");
            worksheet.AddNextFormattedTextCell(text1);
            worksheet.AddNextFormattedTextCell(text2);
            Cell cell1 = worksheet.GetCell(0, 0);
            Cell cell2 = worksheet.GetCell(1, 0);
            Assert.NotNull(cell1);
            Assert.NotNull(cell2);
        }

        [Fact(DisplayName = "Test of AddFormattedTextCellRange with Address objects")]
        public void AddFormattedTextCellRangeWithAddressObjectsTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            List<FormattedText> texts = new List<FormattedText>
            {
                new FormattedText().AddRun("Text1"),
                new FormattedText().AddRun("Text2"),
                new FormattedText().AddRun("Text3")
            };
            Address startAddress = new Address(0, 0);
            Address endAddress = new Address(2, 0);
            worksheet.AddFormattedTextCellRange(texts, startAddress, endAddress);
            Cell cell1 = worksheet.GetCell(0, 0);
            Cell cell2 = worksheet.GetCell(1, 0);
            Cell cell3 = worksheet.GetCell(2, 0);
            Assert.NotNull(cell1);
            Assert.NotNull(cell2);
            Assert.NotNull(cell3);
        }

        [Fact(DisplayName = "Test of AddFormattedTextCellRange with Address as string")]
        public void AddFormattedTextCellRangeWithAddressStringsTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            List<FormattedText> texts = new List<FormattedText>
            {
                new FormattedText().AddRun("Text1"),
                new FormattedText().AddRun("Text2"),
                new FormattedText().AddRun("Text3")
            };
            worksheet.AddFormattedTextCellRange(texts, "A1", "C1");
            Cell cell1 = worksheet.Cells["A1"];
            Cell cell2 = worksheet.Cells["B1"];
            Cell cell3 = worksheet.Cells["C1"];
            Assert.NotNull(cell1);
            Assert.NotNull(cell2);
            Assert.NotNull(cell3);
        }

        [Fact(DisplayName = "Test of AddFormattedTextCellRange with Address objects and style")]
        public void AddFormattedTextCellRangeWithAddressObjectsAndStyleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            List<FormattedText> texts = new List<FormattedText>
            {
                new FormattedText().AddRun("Text1"),
                new FormattedText().AddRun("Text2")
            };
            Address startAddress = new Address(0, 0);
            Address endAddress = new Address(1, 0);
            Style style = new Style();
            style.CurrentFont.Bold = true;
            worksheet.AddFormattedTextCellRange(texts, startAddress, endAddress, style);
            Cell cell1 = worksheet.GetCell(0, 0);
            Cell cell2 = worksheet.GetCell(1, 0);
            Assert.NotNull(cell1.CellStyle);
            Assert.NotNull(cell2.CellStyle);
        }

        [Fact(DisplayName = "Test of AddFormattedTextCellRange with Address as string and style")]
        public void AddFormattedTextCellRangeWithAddressStringsAndStyleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            List<FormattedText> texts = new List<FormattedText>
            {
                new FormattedText().AddRun("Text1"),
                new FormattedText().AddRun("Text2"),
                new FormattedText().AddRun("Text3")
            };
            Style style = new Style();
            style.CurrentFont.Italic = true;
            worksheet.AddFormattedTextCellRange(texts, "B2", "B4", style);
            Cell cell1 = worksheet.Cells["B2"];
            Cell cell2 = worksheet.Cells["B3"];
            Cell cell3 = worksheet.Cells["B4"];
            Assert.NotNull(cell1);
            Assert.NotNull(cell2);
            Assert.NotNull(cell3);
            Assert.NotNull(cell1.CellStyle);
            Assert.NotNull(cell2.CellStyle);
            Assert.NotNull(cell3.CellStyle);
        }

        [Fact(DisplayName = "Test of AddFormattedTextCellRange with string range")]
        public void AddFormattedTextCellRangeWithStringRangeTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            List<FormattedText> texts = new List<FormattedText>
            {
                new FormattedText().AddRun("Text1"),
                new FormattedText().AddRun("Text2"),
                new FormattedText().AddRun("Text3")
            };
            Style style = new Style();
            worksheet.AddFormattedTextCellRange(texts, "A1:C1", style);
            Cell cell1 = worksheet.GetCell(0, 0);
            Cell cell2 = worksheet.GetCell(1, 0);
            Cell cell3 = worksheet.GetCell(2, 0);
            Assert.NotNull(cell1);
            Assert.NotNull(cell2);
            Assert.NotNull(cell3);
            Assert.Equal("Text1", cell1.Value.ToString());
            Assert.Equal("Text2", cell2.Value.ToString());
            Assert.Equal("Text3", cell3.Value.ToString());
        }

        [Fact(DisplayName = "Test of AddFormattedTextCellRange with line breaks but no style")]
        public void AddFormattedTextCellRangeWithLineBreakAndNoStyleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText text1 = new FormattedText();
            text1.AddRun("Text1").AddLineBreak();
            FormattedText text2 = new FormattedText();
            text2.AddRun("Text2");
            List<FormattedText> texts = new List<FormattedText>
            {
                text1,
                text2
            };
            Address startAddress = new Address("D7");
            Address endAddress = new Address("D8");
            worksheet.AddFormattedTextCellRange(texts, startAddress, endAddress);
            Cell cell1 = worksheet.GetCell(3, 6);
            Cell cell2 = worksheet.GetCell(3, 7);
            Assert.NotNull(cell1.CellStyle);
            Assert.Null(cell2.CellStyle);
            Assert.Equal(FormattedText.LineBreakStyle.GetHashCode(), cell1.CellStyle.GetHashCode());
        }

        [Theory(DisplayName = "Test of AddFormattedTextCellRange with various string ranges")]
        [InlineData("A1:B2")]
        [InlineData("$A$1:$B$2")]
        [InlineData("a1:b2")]
        public void AddFormattedTextCellRangeWithVariousStringRangesTest(string range)
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            List<FormattedText> texts = new List<FormattedText>
            {
                new FormattedText().AddRun("1"),
                new FormattedText().AddRun("2"),
                new FormattedText().AddRun("3"),
                new FormattedText().AddRun("4")
            };
            Style style = new Style();
            worksheet.AddFormattedTextCellRange(texts, range, style);
            Cell cell = worksheet.GetCell(0, 0);
            Assert.NotNull(cell);
        }

        [Fact(DisplayName = "Test of AddFormattedTextCellRange with null list throws WorksheetException")]
        public void AddFormattedTextCellRangeWithNullListTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            Address startAddress = new Address(0, 0);
            Address endAddress = new Address(1, 0);
            Assert.Throws<WorksheetException>(() => worksheet.AddFormattedTextCellRange(null, startAddress, endAddress));
        }

        [Fact(DisplayName = "Test of AddFormattedTextCellRange with empty list throws WorksheetException")]
        public void AddFormattedTextCellRangeWithEmptyListTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            List<FormattedText> texts = new List<FormattedText>();
            Address startAddress = new Address(0, 0);
            Address endAddress = new Address(1, 0);
            Style style = new Style();
            Assert.Throws<WorksheetException>(() => worksheet.AddFormattedTextCellRange(texts, startAddress, endAddress, style));
        }

        [Fact(DisplayName = "Test of AddFormattedTextCellRange with null values inserts empty cells")]
        public void AddFormattedTextCellRangeWithNullValuesTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            List<FormattedText> texts = new List<FormattedText>
            {
                new FormattedText().AddRun("Text1"),
                null,
                new FormattedText().AddRun("Text3")
            };
            Address startAddress = new Address(0, 0);
            Address endAddress = new Address(2, 0);
            worksheet.AddFormattedTextCellRange(texts, startAddress, endAddress);
            Cell cell1 = worksheet.GetCell(0, 0);
            Cell cell2 = worksheet.GetCell(1, 0);
            Cell cell3 = worksheet.GetCell(2, 0);
            Assert.NotNull(cell1);
            Assert.NotNull(cell2);
            Assert.NotNull(cell3);
            Assert.Equal(Cell.CellType.Empty, cell2.DataType);
        }

        [Fact(DisplayName = "Test of AddFormattedTextCellRange with WrapText merges styles")]
        public void AddFormattedTextCellRangeWithWrapTextMergesStylesTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText text1 = new FormattedText();
            text1.AddRun("Line1");
            text1.AddLineBreak();
            List<FormattedText> texts = new List<FormattedText> { text1 };
            Address startAddress = new Address(0, 0);
            Address endAddress = new Address(0, 0);
            Style style = new Style();
            style.CurrentFont.Bold = true;
            worksheet.AddFormattedTextCellRange(texts, startAddress, endAddress, style);
            Cell cell = worksheet.GetCell(0, 0);
            Assert.NotNull(cell.CellStyle);
            Assert.True(cell.CellStyle.CurrentFont.Bold);
            Assert.Equal(CellXf.TextBreakValue.WrapText, cell.CellStyle.CurrentCellXf.Alignment);
        }

        [Fact(DisplayName = "Test of AddFormattedTextCellRange with mixed WrapText values")]
        public void AddFormattedTextCellRangeWithMixedWrapTextValuesTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText text1 = new FormattedText();
            text1.AddRun("Wrapped");
            text1.AddLineBreak();
            FormattedText text2 = new FormattedText();
            text2.AddRun("Normal");
            List<FormattedText> texts = new List<FormattedText> { text1, text2 };
            Address startAddress = new Address(0, 0);
            Address endAddress = new Address(1, 0);
            Style style = new Style();
            style.CurrentFont.Italic = true;
            worksheet.AddFormattedTextCellRange(texts, startAddress, endAddress, style);
            Cell cell1 = worksheet.GetCell(0, 0);
            Cell cell2 = worksheet.GetCell(1, 0);
            Assert.Equal(CellXf.TextBreakValue.WrapText, cell1.CellStyle.CurrentCellXf.Alignment);
            Assert.True(cell1.CellStyle.CurrentFont.Italic);
            Assert.True(cell2.CellStyle.CurrentFont.Italic);
        }

        [Fact(DisplayName = "Test of AddFormattedTextCellRange with null FormattedText in list")]
        public void AddFormattedTextCellRangeWithNullFormattedTextInListTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            List<FormattedText> texts = new List<FormattedText>
    {
        new FormattedText().AddRun("Text1"),
        null,
        new FormattedText().AddRun("Text3")
    };
            Address startAddress = new Address(0, 0);
            Address endAddress = new Address(2, 0);
            Style style = new Style();
            style.CurrentFont.Bold = true;

            worksheet.AddFormattedTextCellRange(texts, startAddress, endAddress, style);

            Cell cell1 = worksheet.GetCell(0, 0);
            Cell cell2 = worksheet.GetCell(1, 0);
            Cell cell3 = worksheet.GetCell(2, 0);
            Assert.NotNull(cell1);
            Assert.NotNull(cell2);
            Assert.NotNull(cell3);
            Assert.Equal(Cell.CellType.Empty, cell2.DataType);
            Assert.NotNull(cell2.CellStyle);
        }

        [Fact(DisplayName = "Test of AddFormattedTextCellRange with empty PlainText in list")]
        public void AddFormattedTextCellRangeWithEmptyPlainTextInListTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText emptyText = new FormattedText();
            List<FormattedText> texts = new List<FormattedText>
    {
        new FormattedText().AddRun("Text1"),
        emptyText,
        new FormattedText().AddRun("Text3")
    };
            Address startAddress = new Address(0, 0);
            Address endAddress = new Address(2, 0);
            Style style = new Style();
            style.CurrentFont.Italic = true;

            worksheet.AddFormattedTextCellRange(texts, startAddress, endAddress, style);

            Cell cell1 = worksheet.GetCell(0, 0);
            Cell cell2 = worksheet.GetCell(1, 0);
            Cell cell3 = worksheet.GetCell(2, 0);
            Assert.NotNull(cell2);
            Assert.Equal(Cell.CellType.Empty, cell2.DataType);
            Assert.NotNull(cell2.CellStyle);
        }

        [Fact(DisplayName = "Test of AddFormattedTextCellRange with null in list and no style")]
        public void AddFormattedTextCellRangeWithNullInListAndNoStyleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            List<FormattedText> texts = new List<FormattedText>
    {
        new FormattedText().AddRun("Text1"),
        null,
        new FormattedText().AddRun("Text3")
    };
            Address startAddress = new Address(0, 0);
            Address endAddress = new Address(2, 0);

            worksheet.AddFormattedTextCellRange(texts, startAddress, endAddress);

            Cell cell2 = worksheet.GetCell(1, 0);
            Assert.NotNull(cell2);
            Assert.Equal(Cell.CellType.Empty, cell2.DataType);
        }

        [Fact(DisplayName = "Test of AddFormattedTextCellRange with WrapText true and no style")]
        public void AddFormattedTextCellRangeWithWrapTextAndNoStyleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText text1 = new FormattedText();
            text1.AddRun("Line1");
            text1.AddLineBreak();
            text1.AddRun("Line2");
            FormattedText text2 = new FormattedText();
            text2.AddRun("Normal");
            List<FormattedText> texts = new List<FormattedText> { text1, text2 };
            Address startAddress = new Address(0, 0);
            Address endAddress = new Address(1, 0);

            worksheet.AddFormattedTextCellRange(texts, startAddress, endAddress);

            Cell cell1 = worksheet.GetCell(0, 0);
            Cell cell2 = worksheet.GetCell(1, 0);
            Assert.NotNull(cell1);
            Assert.NotNull(cell1.CellStyle);
            Assert.Equal(CellXf.TextBreakValue.WrapText, cell1.CellStyle.CurrentCellXf.Alignment);
            Assert.NotNull(cell2);
        }

        [Fact(DisplayName = "Test of AddFormattedTextCellRange with no WrapText and no style")]
        public void AddFormattedTextCellRangeWithNoWrapTextAndNoStyleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText text1 = new FormattedText();
            text1.AddRun("Text1");
            FormattedText text2 = new FormattedText();
            text2.AddRun("Text2");
            List<FormattedText> texts = new List<FormattedText> { text1, text2 };
            Address startAddress = new Address(0, 0);
            Address endAddress = new Address(1, 0);

            worksheet.AddFormattedTextCellRange(texts, startAddress, endAddress);

            Cell cell1 = worksheet.GetCell(0, 0);
            Cell cell2 = worksheet.GetCell(1, 0);
            Assert.NotNull(cell1);
            Assert.NotNull(cell2);
        }

        [Fact(DisplayName = "Test of AddFormattedTextCellRange with all null values and style")]
        public void AddFormattedTextCellRangeWithAllNullValuesAndStyleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            List<FormattedText> texts = new List<FormattedText> { null, null, null };
            Address startAddress = new Address(0, 0);
            Address endAddress = new Address(2, 0);
            Style style = new Style();
            style.CurrentFont.Bold = true;

            worksheet.AddFormattedTextCellRange(texts, startAddress, endAddress, style);

            Cell cell1 = worksheet.GetCell(0, 0);
            Cell cell2 = worksheet.GetCell(1, 0);
            Cell cell3 = worksheet.GetCell(2, 0);
            Assert.Equal(Cell.CellType.Empty, cell1.DataType);
            Assert.Equal(Cell.CellType.Empty, cell2.DataType);
            Assert.Equal(Cell.CellType.Empty, cell3.DataType);
            Assert.NotNull(cell1.CellStyle);
            Assert.NotNull(cell2.CellStyle);
            Assert.NotNull(cell3.CellStyle);
        }

        [Fact(DisplayName = "Test of AddFormattedTextCellRange with all empty PlainText and no style")]
        public void AddFormattedTextCellRangeWithAllEmptyPlainTextAndNoStyleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText empty1 = new FormattedText();
            FormattedText empty2 = new FormattedText();
            List<FormattedText> texts = new List<FormattedText> { empty1, empty2 };
            Address startAddress = new Address(0, 0);
            Address endAddress = new Address(1, 0);

            worksheet.AddFormattedTextCellRange(texts, startAddress, endAddress);

            Cell cell1 = worksheet.GetCell(0, 0);
            Cell cell2 = worksheet.GetCell(1, 0);
            Assert.Equal(Cell.CellType.Empty, cell1.DataType);
            Assert.Null(cell1.Value);
            Assert.Equal(Cell.CellType.Empty, cell2.DataType);
            Assert.Null(cell2.Value);
            workbook.SaveAs(@"c:\purge-temp\zomgue01.xlsx");
        }

        [Fact(DisplayName = "Test of AddFormattedTextCellRange with mixed null, empty and valid values without style")]
        public void AddFormattedTextCellRangeWithMixedValuesNoStyleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText emptyText = new FormattedText();
            FormattedText validText = new FormattedText().AddRun("Valid");
            List<FormattedText> texts = new List<FormattedText> { null, emptyText, validText };
            Address startAddress = new Address(0, 0);
            Address endAddress = new Address(2, 0);

            worksheet.AddFormattedTextCellRange(texts, startAddress, endAddress);

            Cell cell1 = worksheet.GetCell(0, 0);
            Cell cell2 = worksheet.GetCell(1, 0);
            Cell cell3 = worksheet.GetCell(2, 0);
            Assert.Equal(Cell.CellType.Empty, cell1.DataType);
            Assert.Equal(Cell.CellType.Empty, cell2.DataType);
            Assert.Equal(Cell.CellType.String, cell3.DataType);
            Assert.Null(cell1.Value);
            Assert.Null(cell2.Value);
            Assert.IsType<FormattedText>(cell3.Value);
            Assert.Equal("Valid", cell3.Value.ToString());
        }

        [Fact(DisplayName = "Test of AddFormattedTextCellRange with WrapText false and style")]
        public void AddFormattedTextCellRangeWithNoWrapTextAndStyleTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText text1 = new FormattedText();
            text1.AddRun("Text1");
            FormattedText text2 = new FormattedText();
            text2.AddRun("Text2");
            List<FormattedText> texts = new List<FormattedText> { text1, text2 };
            Address startAddress = new Address(0, 0);
            Address endAddress = new Address(1, 0);
            Style style = new Style();
            style.CurrentFont.Underline = Font.UnderlineValue.Single;

            worksheet.AddFormattedTextCellRange(texts, startAddress, endAddress, style);

            Cell cell1 = worksheet.GetCell(0, 0);
            Cell cell2 = worksheet.GetCell(1, 0);
            Assert.NotNull(cell1);
            Assert.NotNull(cell2);
            Assert.NotNull(cell1.CellStyle);
            Assert.NotNull(cell2.CellStyle);
            Assert.Equal(Font.UnderlineValue.Single, cell1.CellStyle.CurrentFont.Underline);
        }

        [Fact(DisplayName = "Test of AddFormattedTextCell does not modify original FormattedText")]
        public void AddFormattedTextCellDoesNotModifyOriginalTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            FormattedText formattedText = new FormattedText();
            formattedText.AddRun("Test");
            string originalText = formattedText.PlainText;
            worksheet.AddFormattedTextCell(formattedText, 0, 0);
            Assert.Equal(originalText, formattedText.PlainText);
        }
    }
}
