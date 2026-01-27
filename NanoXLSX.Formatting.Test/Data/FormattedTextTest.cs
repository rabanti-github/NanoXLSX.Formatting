using NanoXLSX.Extensions;
using NanoXLSX.Styles;
using System;
using Xunit;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX.Tests
{
    public class FormattedTextTests
    {
        [Fact(DisplayName = "Test of the default constructor")]
        public void ConstructorTest()
        {
            FormattedText text = new FormattedText();

            Assert.NotNull(text);
            Assert.Empty(text.Runs);
            Assert.Empty(text.PhoneticRuns);
            Assert.Null(text.PhoneticProperties);
            Assert.False(text.WrapText);
            Assert.Empty(text.PlainText);
        }

        [Theory(DisplayName = "Test of the WrapText property")]
        [InlineData(false)]
        [InlineData(true)]
        public void WrapTextPropertyTest(bool expectedValue)
        {
            FormattedText text = new FormattedText();
            Assert.False(text.WrapText);
            text.WrapText = expectedValue;
            Assert.Equal(expectedValue, text.WrapText);
        }

        [Fact(DisplayName = "Test of the Runs property getter")]
        public void RunsPropertyTest()
        {
            FormattedText text = new FormattedText();
            Assert.NotNull(text.Runs);
            Assert.Empty(text.Runs);
            text.AddRun("Test");
            Assert.Single(text.Runs);
            text.AddRun("Another");
            Assert.Equal(2, text.Runs.Count);
        }

        [Fact(DisplayName = "Test of the Runs property returns readonly list")]
        public void RunsPropertyReturnsReadonlyListTest()
        {
            FormattedText text = new FormattedText();
            Assert.IsAssignableFrom<System.Collections.Generic.IReadOnlyList<TextRun>>(text.Runs);
        }

        [Fact(DisplayName = "Test of the PhoneticRuns property getter")]
        public void PhoneticRunsPropertyTest()
        {
            FormattedText text = new FormattedText();

            Assert.NotNull(text.PhoneticRuns);
            Assert.Empty(text.PhoneticRuns);

            text.AddPhoneticRun("テスト", 0, 2);
            Assert.Single(text.PhoneticRuns);
            text.AddPhoneticRun("テスト二", 3, 5);
            Assert.Equal(2, text.PhoneticRuns.Count);
        }

        [Fact(DisplayName = "Test of the PhoneticRuns property returns readonly list")]
        public void PhoneticRunsPropertyReturnsReadonlyListTest()
        {
            FormattedText text = new FormattedText();
            Assert.IsAssignableFrom<System.Collections.Generic.IReadOnlyList<PhoneticRun>>(text.PhoneticRuns);
        }

        [Fact(DisplayName = "Test of the PhoneticProperties property getter and setter")]
        public void PhoneticPropertiesPropertyTest()
        {
            FormattedText text = new FormattedText();
            Font font = new Font();
            font.Italic = true;
            PhoneticProperties properties = new PhoneticProperties(font);

            Assert.Null(text.PhoneticProperties);
            text.PhoneticProperties = properties;
            Assert.Equal(properties, text.PhoneticProperties);
            Assert.True(text.PhoneticProperties.FontReference.Italic);
        }

        [Fact(DisplayName = "Test of the PhoneticProperties property setter with null")]
        public void PhoneticPropertiesPropertySetNullTest()
        {
            FormattedText text = new FormattedText();
            Font font = new Font();
            text.PhoneticProperties = new PhoneticProperties(font);
            text.PhoneticProperties = null;
            Assert.Null(text.PhoneticProperties);
        }

        [Theory(DisplayName = "Test of the PlainText property with single run")]
        [InlineData("Hello")]
        [InlineData("テスト")]
        [InlineData("你好")]
        [InlineData("　　")]
        [InlineData("0123")]
        [InlineData("\t")]
        [InlineData("\n\n")]
        [InlineData("")]
        public void PlainTextPropertySingleRunTest(string expectedText)
        {
            FormattedText text = new FormattedText();
            text.AddRun(expectedText == "" ? " " : expectedText);
            Assert.Equal(expectedText == "" ? " " : expectedText, text.PlainText);
        }

        [Fact(DisplayName = "Test of the PlainText property with multiple runs")]
        public void PlainTextPropertyMultipleRunsTest()
        {
            FormattedText text = new FormattedText();
            text.AddRun("_Run1_");
            text.AddRun(" ");
            text.AddRun("\n");
            text.AddRun("テスト");
            Assert.Equal("_Run1_ \nテスト", text.PlainText);
        }

        [Fact(DisplayName = "Test of the PlainText property with empty FormattedText")]
        public void PlainTextPropertyEmptyTest()
        {
            FormattedText text = new FormattedText();
            Assert.Empty(text.PlainText);
        }

        [Theory(DisplayName = "Test of the AddRun method with text only")]
        [InlineData("Test")]
        [InlineData("Hello World")]
        [InlineData("123")]
        [InlineData("テスト")]
        [InlineData("你好")]
        [InlineData("　　")]
        [InlineData("\t")]
        [InlineData("\r\n")]
        public void AddRunWithTextOnlyTest(string runText)
        {
            FormattedText text = new FormattedText();
            FormattedText result = text.AddRun(runText);
            Assert.Same(text, result);
            Assert.Single(text.Runs);
            Assert.Equal(runText, text.PlainText);
        }

        [Fact(DisplayName = "Test of the AddRun method with style builder")]
        public void AddRunWithStyleBuilderTest()
        {
            FormattedText text = new FormattedText();
            FormattedText result = text.AddRun("Test", builder => builder.Bold());

            Assert.Same(text, result);
            Assert.Single(text.Runs);
            Assert.Equal("Test", text.PlainText);
            Assert.NotNull(text.Runs[0].FontStyle);
            Assert.True(text.Runs[0].FontStyle.Bold);
        }

        [Fact(DisplayName = "Test of the AddRun method with style builder and null action")]
        public void AddRunWithNullStyleBuilderActionTest()
        {
            FormattedText text = new FormattedText();
            FormattedText result = text.AddRun("Test", (Action<InlineStyleBuilder>)null);

            Assert.Same(text, result);
            Assert.Single(text.Runs);
            Assert.Equal("Test", text.PlainText);
        }

        [Fact(DisplayName = "Test of the AddRun method with style builder and multiple style operations")]
        public void AddRunWithStyleBuilderMultipleOperationsTest()
        {
            FormattedText text = new FormattedText();
            FormattedText result = text.AddRun("Test", builder =>
            {
                builder.Bold();
                builder.Italic();
                builder.Size(14);
            });

            Assert.Same(text, result);
            Assert.Single(text.Runs);
            Assert.NotNull(text.Runs[0].FontStyle);
            Assert.True(text.Runs[0].FontStyle.Bold);
            Assert.True(text.Runs[0].FontStyle.Italic);
            Assert.Equal(14, text.Runs[0].FontStyle.Size);
        }

        [Fact(DisplayName = "Test of the AddRun method with style builder and null text - should throw FormatException")]
        public void AddRunWithStyleBuilderAndNullTextTest()
        {
            FormattedText text = new FormattedText();
            Assert.Throws<FormatException>(() => text.AddRun(null, builder => builder.Bold()));
        }

        [Fact(DisplayName = "Test of the AddRun method with style builder and empty text - should throw FormatException")]
        public void AddRunWithStyleBuilderAndEmptyTextTest()
        {
            FormattedText text = new FormattedText();
            Assert.Throws<FormatException>(() => text.AddRun("", builder => builder.Bold()));
        }

        [Fact(DisplayName = "Test of the AddRun method with style builder chaining")]
        public void AddRunWithStyleBuilderChainingTest()
        {
            FormattedText text = new FormattedText();
            FormattedText result = text
                .AddRun("Bold", builder => builder.Bold())
                .AddRun(" ")
                .AddRun("Italic", builder => builder.Italic());

            Assert.Same(text, result);
            Assert.Equal(3, text.Runs.Count);
            Assert.Equal("Bold Italic", text.PlainText);
        }

        [Fact(DisplayName = "Test of the AddRun method with style builder containing line breaks sets WrapText")]
        public void AddRunWithStyleBuilderAndLineBreakSetsWrapTextTest()
        {
            FormattedText text = new FormattedText();
            text.AddRun("Line1\nLine2", builder => builder.Bold());

            Assert.True(text.WrapText);
        }

        [Fact(DisplayName = "Test of the AddRun method with text and font style")]
        public void AddRunWithTextAndFontTest()
        {
            FormattedText text = new FormattedText();
            Font font = new Font();
            font.Bold = true;
            FormattedText result = text.AddRun("Test", font);
            Assert.Same(text, result);
            Assert.Single(text.Runs);
            Assert.Equal("Test", text.PlainText);
            Assert.NotNull(text.Runs[0].FontStyle);
            Assert.True(text.Runs[0].FontStyle.Bold);
        }

        [Theory(DisplayName = "Test of the AddRun method with null text - should throw FormatException")]
        [InlineData(null)]
        public void AddRunWithNullTextFailingTest(string runText)
        {
            FormattedText text = new FormattedText();
            Assert.Throws<FormatException>(() => text.AddRun(runText));
        }

        [Fact(DisplayName = "Test of the AddRun method with empty text - should throw FormatException")]
        public void AddRunWithEmptyTextFailingTest()
        {
            FormattedText text = new FormattedText();
            Assert.Throws<FormatException>(() => text.AddRun(""));
        }

        [Fact(DisplayName = "Test of the AddRun method with null text and font - should throw FormatException")]
        public void AddRunWithNullTextAndFontFailingTest()
        {
            FormattedText text = new FormattedText();
            Font font = new Font();
            Assert.Throws<FormatException>(() => text.AddRun(null, font));
        }

        [Fact(DisplayName = "Test of the AddRun method with empty text and font - should throw FormatException")]
        public void AddRunWithEmptyTextAndFontFailingTest()
        {
            FormattedText text = new FormattedText();
            Font font = new Font();
            Assert.Throws<FormatException>(() => text.AddRun("", font));
        }

        [Fact(DisplayName = "Test of the AddRun method chaining")]
        public void AddRunChainingTest()
        {
            FormattedText text = new FormattedText();
            FormattedText result = text.AddRun("Hello").AddRun(" ").AddRun("World");
            Assert.Same(text, result);
            Assert.Equal(3, text.Runs.Count);
            Assert.Equal("Hello World", text.PlainText);
        }

        [Fact(DisplayName = "Test of the AddRun method with line break in text sets WrapText")]
        public void AddRunWithLineBreakSetsWrapTextTest()
        {
            FormattedText text = new FormattedText();
            text.AddRun("Line1\nLine2");
            Assert.True(text.WrapText);
        }

        [Fact(DisplayName = "Test of the AddRun method with carriage return and line feed sets WrapText")]
        public void AddRunWithCarriageReturnLineFeedSetsWrapTextTest()
        {
            FormattedText text = new FormattedText();
            text.AddRun("Line1\r\nLine2");
            Assert.True(text.WrapText);
        }

        [Theory(DisplayName = "Test of the AddLineBreak method with useStyleFromLastRun parameter")]
        [InlineData(true)]
        [InlineData(false)]
        public void AddLineBreakWithParameterTest(bool useStyleFromLastRun)
        {
            FormattedText text = new FormattedText();
            text.AddRun("First line");
            int initialRunCount = text.Runs.Count;
            text.AddLineBreak(useStyleFromLastRun);
            Assert.True(text.WrapText);
            if (useStyleFromLastRun)
            {
                Assert.Equal(initialRunCount, text.Runs.Count);
                Assert.Contains(Environment.NewLine, text.PlainText);
            }
            else
            {
                Assert.Equal(initialRunCount + 1, text.Runs.Count);
            }
        }

        [Fact(DisplayName = "Test of the AddLineBreak method on empty FormattedText")]
        public void AddLineBreakOnEmptyFormattedTextTest()
        {
            FormattedText text = new FormattedText();
            text.AddLineBreak();
            Assert.True(text.WrapText);
            Assert.Single(text.Runs);
            Assert.Equal(Environment.NewLine, text.PlainText);
        }

        [Fact(DisplayName = "Test of the AddLineBreak method on empty FormattedText with useStyleFromLastRun false")]
        public void AddLineBreakOnEmptyFormattedTextWithFalseParameterTest()
        {
            FormattedText text = new FormattedText();
            text.AddLineBreak(false);
            Assert.True(text.WrapText);
            Assert.Single(text.Runs);
            Assert.Equal(Environment.NewLine, text.PlainText);
        }

        [Fact(DisplayName = "Test of the AddLineBreak method sets WrapText to true")]
        public void AddLineBreakSetsWrapTextTest()
        {
            FormattedText text = new FormattedText();
            text.AddRun("Test");
            Assert.False(text.WrapText);
            text.AddLineBreak();
            Assert.True(text.WrapText);
        }

        [Fact(DisplayName = "Test of the AddLineBreak method appends to last run when useStyleFromLastRun is true")]
        public void AddLineBreakAppendsToLastRunTest()
        {
            FormattedText text = new FormattedText();
            text.AddRun("First");
            string expectedText = "First" + Environment.NewLine;
            text.AddLineBreak(true);
            Assert.Single(text.Runs);
            Assert.Equal(expectedText, text.Runs[0].Text);
            Assert.Equal(expectedText, text.PlainText);
        }

        [Fact(DisplayName = "Test of the AddLineBreak method creates new run when useStyleFromLastRun is false")]
        public void AddLineBreakCreatesNewRunTest()
        {
            FormattedText text = new FormattedText();
            Font font = new Font();
            font.Bold = true;
            text.AddRun("First", font);
            text.AddLineBreak(false);
            Assert.Equal(2, text.Runs.Count);
            Assert.Equal("First", text.Runs[0].Text);
            Assert.Equal(Environment.NewLine, text.Runs[1].Text);
            Assert.NotNull(text.Runs[0].FontStyle);
            Assert.Null(text.Runs[1].FontStyle);
        }

        [Fact(DisplayName = "Test of the AddLineBreak method default parameter")]
        public void AddLineBreakDefaultParameterTest()
        {
            FormattedText text = new FormattedText();
            text.AddRun("Test");
            int initialRunCount = text.Runs.Count;
            text.AddLineBreak();
            Assert.Equal(initialRunCount, text.Runs.Count);
            Assert.Contains(Environment.NewLine, text.PlainText);
        }

        [Theory(DisplayName = "Test of the AddPhoneticRun method")]
        [InlineData("テスト", 0u, 2u)]
        [InlineData("你好", 5u, 10u)]
        [InlineData("ふりがな", 0u, 0u)]
        [InlineData("nǐ hǎo", 0u, 1u)]
        public void AddPhoneticRunTest(string phoneticText, uint startBase, uint endBase)
        {
            FormattedText text = new FormattedText();
            FormattedText result = text.AddPhoneticRun(phoneticText, startBase, endBase);
            Assert.Same(text, result);
            Assert.Single(text.PhoneticRuns);
            Assert.Equal(phoneticText, text.PhoneticRuns[0].Text);
            Assert.Equal(startBase, text.PhoneticRuns[0].StartBase);
            Assert.Equal(endBase, text.PhoneticRuns[0].EndBase);
        }

        [Fact(DisplayName = "Test of the AddPhoneticRun method with null text - should throw FormatException")]
        public void AddPhoneticRunWithNullTextFailingTest()
        {
            FormattedText text = new FormattedText();
            Assert.Throws<FormatException>(() => text.AddPhoneticRun(null, 0, 2));
        }

        [Fact(DisplayName = "Test of the AddPhoneticRun method with empty text - should throw FormatException")]
        public void AddPhoneticRunWithEmptyTextFailingTest()
        {
            FormattedText text = new FormattedText();
            Assert.Throws<FormatException>(() => text.AddPhoneticRun("", 0, 2));
        }

        [Fact(DisplayName = "Test of the AddPhoneticRun method chaining")]
        public void AddPhoneticRunChainingTest()
        {
            FormattedText text = new FormattedText();
            FormattedText result = text.AddPhoneticRun("テスト", 0, 2)
                                        .AddPhoneticRun("ピンイン", 3, 5);

            Assert.Same(text, result);
            Assert.Equal(2, text.PhoneticRuns.Count);
        }

        [Theory(DisplayName = "Test of the SetPhoneticProperties method with different parameters")]
        [InlineData(PhoneticRun.PhoneticType.FullwidthKatakana, PhoneticRun.PhoneticAlignment.Left)]
        [InlineData(PhoneticRun.PhoneticType.HalfwidthKatakana, PhoneticRun.PhoneticAlignment.Center)]
        [InlineData(PhoneticRun.PhoneticType.Hiragana, PhoneticRun.PhoneticAlignment.Distributed)]
        [InlineData(PhoneticRun.PhoneticType.NoConversion, PhoneticRun.PhoneticAlignment.NoControl)]
        public void SetPhoneticPropertiesTest(PhoneticRun.PhoneticType type, PhoneticRun.PhoneticAlignment alignment)
        {
            FormattedText text = new FormattedText();
            Font font = new Font();
            FormattedText result = text.SetPhoneticProperties(font, type, alignment);
            Assert.Same(text, result);
            Assert.NotNull(text.PhoneticProperties);
            Assert.Equal(type, text.PhoneticProperties.Type);
            Assert.Equal(alignment, text.PhoneticProperties.Alignment);
        }

        [Fact(DisplayName = "Test of the SetPhoneticProperties method with default parameters")]
        public void SetPhoneticPropertiesWithDefaultsTest()
        {
            FormattedText text = new FormattedText();
            Font font = new Font();
            FormattedText result = text.SetPhoneticProperties(font);
            Assert.Same(text, result);
            Assert.NotNull(text.PhoneticProperties);
            Assert.Equal(PhoneticRun.PhoneticType.FullwidthKatakana, text.PhoneticProperties.Type);
            Assert.Equal(PhoneticRun.PhoneticAlignment.Left, text.PhoneticProperties.Alignment);
        }

        [Fact(DisplayName = "Test of the SetPhoneticProperties method overwrites existing properties")]
        public void SetPhoneticPropertiesOverwritesTest()
        {
            FormattedText text = new FormattedText();
            Font font1 = new Font();
            Font font2 = new Font();
            font2.Bold = true;
            text.SetPhoneticProperties(font1, PhoneticRun.PhoneticType.Hiragana, PhoneticRun.PhoneticAlignment.Center);
            text.SetPhoneticProperties(font2, PhoneticRun.PhoneticType.FullwidthKatakana, PhoneticRun.PhoneticAlignment.Left);
            Assert.Equal(PhoneticRun.PhoneticType.FullwidthKatakana, text.PhoneticProperties.Type);
            Assert.Equal(PhoneticRun.PhoneticAlignment.Left, text.PhoneticProperties.Alignment);
        }

        [Fact(DisplayName = "Test of the Clear method")]
        public void ClearTest()
        {
            FormattedText text = new FormattedText();
            Font font = new Font();
            text.AddRun("Test");
            text.AddPhoneticRun("テスト", 0, 2);
            text.SetPhoneticProperties(font);
            text.Clear();
            Assert.Empty(text.Runs);
            Assert.Empty(text.PhoneticRuns);
            Assert.Null(text.PhoneticProperties);
        }

        [Fact(DisplayName = "Test of the Clear method on empty FormattedText")]
        public void ClearOnEmptyFormattedTextTest()
        {
            FormattedText text = new FormattedText();
            text.Clear();
            Assert.Empty(text.Runs);
            Assert.Empty(text.PhoneticRuns);
            Assert.Null(text.PhoneticProperties);
        }

        [Fact(DisplayName = "Test of the Clear method with multiple runs and phonetic runs")]
        public void ClearWithMultipleElementsTest()
        {
            FormattedText text = new FormattedText();
            Font font = new Font();
            text.AddRun("First");
            text.AddRun("Second");
            text.AddRun("Third");
            text.AddPhoneticRun("テスト1", 0, 2);
            text.AddPhoneticRun("テスト2", 3, 5);
            text.SetPhoneticProperties(font);
            text.WrapText = true;
            text.Clear();
            Assert.Empty(text.Runs);
            Assert.Empty(text.PhoneticRuns);
            Assert.Null(text.PhoneticProperties);
        }

        [Fact(DisplayName = "Test of the Clear method does not reset WrapText")]
        public void ClearDoesNotResetWrapTextTest()
        {
            FormattedText text = new FormattedText();
            text.AddRun("Test");
            text.WrapText = true;
            text.Clear();
            Assert.True(text.WrapText);
        }

        [Fact(DisplayName = "Test of the Copy method")]
        public void CopyTest()
        {
            FormattedText text = new FormattedText();
            Font font = new Font();
            font.Bold = true;
            text.AddRun("Hello", font);
            text.AddRun("World");
            text.AddPhoneticRun("テスト", 0, 2);
            text.SetPhoneticProperties(font);
            text.WrapText = true;
            FormattedText copy = text.Copy();
            Assert.NotSame(text, copy);
            Assert.Equal(text.Runs.Count, copy.Runs.Count);
            Assert.Equal(text.PhoneticRuns.Count, copy.PhoneticRuns.Count);
            Assert.Equal(text.PlainText, copy.PlainText);
            Assert.NotNull(copy.PhoneticProperties);
        }

        [Fact(DisplayName = "Test of the Copy method on empty FormattedText")]
        public void CopyOnEmptyFormattedTextTest()
        {
            FormattedText text = new FormattedText();
            FormattedText copy = text.Copy();
            Assert.NotSame(text, copy);
            Assert.Empty(copy.Runs);
            Assert.Empty(copy.PhoneticRuns);
            Assert.Null(copy.PhoneticProperties);
        }

        [Fact(DisplayName = "Test of the Copy method creates deep copy of runs")]
        public void CopyCreatesDeepCopyOfRunsTest()
        {
            FormattedText text = new FormattedText();
            text.AddRun("Original");
            FormattedText copy = text.Copy();
            copy.AddRun("Modified");
            Assert.Single(text.Runs);
            Assert.Equal(2, copy.Runs.Count);
            Assert.Equal("Original", text.PlainText);
            Assert.Equal("OriginalModified", copy.PlainText);
        }

        [Fact(DisplayName = "Test of the Copy method creates deep copy of phonetic runs")]
        public void CopyCreatesDeepCopyOfPhoneticRunsTest()
        {
            FormattedText text = new FormattedText();
            text.AddPhoneticRun("テスト", 0, 2);
            FormattedText copy = text.Copy();
            copy.AddPhoneticRun("ピンイン", 3, 5);
            Assert.Single(text.PhoneticRuns);
            Assert.Equal(2, copy.PhoneticRuns.Count);
        }

        [Fact(DisplayName = "Test of the Copy method does not copy WrapText property")]
        public void CopyDoesNotCopyWrapTextTest()
        {
            FormattedText text = new FormattedText();
            text.AddRun("Test");
            text.WrapText = true;
            FormattedText copy = text.Copy();

            Assert.False(copy.WrapText);
        }

        [Theory(DisplayName = "Test of the ToString method")]
        [InlineData("Test")]
        [InlineData("Hello World")]
        [InlineData(" ")]
        [InlineData("你好")]
        [InlineData("ふりがな")]
        [InlineData("nǐ hǎo")]
        [InlineData("\n")]
        public void ToStringTest(string expectedText)
        {
            FormattedText text = new FormattedText();
            text.AddRun(expectedText);
            Assert.Equal(expectedText, text.ToString());
        }

        [Fact(DisplayName = "Test of the ToString method with multiple runs")]
        public void ToStringWithMultipleRunsTest()
        {
            FormattedText text = new FormattedText();
            text.AddRun("Hello");
            text.AddRun(" ");
            text.AddRun("World");
            Assert.Equal("Hello World", text.ToString());
        }

        [Fact(DisplayName = "Test of the ToString method on empty FormattedText")]
        public void ToStringOnEmptyFormattedTextTest()
        {
            FormattedText text = new FormattedText();
            Assert.Empty(text.ToString());
        }

        [Fact(DisplayName = "Test of the ToString method equals PlainText")]
        public void ToStringEqualsPlainTextTest()
        {
            FormattedText text = new FormattedText();
            text.AddRun("Hello");
            text.AddRun(" ");
            text.AddRun("World");

            Assert.Equal(text.PlainText, text.ToString());
        }

        [Fact(DisplayName = "Test of the LineBreakStyle static field")]
        public void LineBreakStyleTest()
        {
            Assert.NotNull(FormattedText.LineBreakStyle);
            Assert.Equal(CellXf.TextBreakValue.WrapText, FormattedText.LineBreakStyle.CurrentCellXf.Alignment);
        }

        [Fact(DisplayName = "Test of the LineBreakStyle static field is initialized once")]
        public void LineBreakStyleIsSingletonTest()
        {
            Style style1 = FormattedText.LineBreakStyle;
            Style style2 = FormattedText.LineBreakStyle;
            Assert.Same(style1, style2);
        }
    }
}