using NanoXLSX.Extensions;
using NanoXLSX.Styles;
using System;
using Xunit;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX.Formatting.Test.Builder
{
    public  class FormattedTextBuilderTest
    {
        [Fact(DisplayName = "Test of the default constructor")]
        public void ConstructorTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            Assert.NotNull(builder);
        }

        [Fact(DisplayName = "Test of the AddRun method with text only")]
        public void AddRunWithTextOnlyTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            FormattedTextBuilder result = builder.AddRun("Test");
            Assert.Same(builder, result);
        }

        [Theory(DisplayName = "Test of the AddRun method with various text values")]
        [InlineData("Hello")]
        [InlineData("World")]
        [InlineData("123")]
        [InlineData("　　")]
        [InlineData("0123")]
        [InlineData("\t")]
        [InlineData("\n\n")]
        public void AddRunWithVariousTextValuesTest(string text)
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            FormattedTextBuilder result = builder.AddRun(text);
            Assert.Same(builder, result);
            FormattedText formattedText = builder.Build();
            Assert.Equal(text, formattedText.PlainText);
        }

        [Fact(DisplayName = "Test of the AddRun method with text and font")]
        public void AddRunWithTextAndFontTest()
        {
            Font font = new Font();
            font.Bold = true;
            FormattedTextBuilder builder = new FormattedTextBuilder();
            FormattedTextBuilder result = builder.AddRun("Test", font);
            Assert.Same(builder, result);
            FormattedText formattedText = builder.Build();
            Assert.Single(formattedText.Runs);
            Assert.NotNull(formattedText.Runs[0].FontStyle);
        }

        [Fact(DisplayName = "Test of the AddRun method with text and null font")]
        public void AddRunWithTextAndNullFontTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            FormattedTextBuilder result = builder.AddRun("Test", (Font)null);
            Assert.Same(builder, result);
            FormattedText formattedText = builder.Build();
            Assert.Single(formattedText.Runs);
            Assert.Null(formattedText.Runs[0].FontStyle);
        }

        [Fact(DisplayName = "Test of the AddRun method with null text - should throw FormatException")]
        public void AddRunWithNullTextFailingTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            Assert.Throws<FormatException>(() => builder.AddRun(null));
        }

        [Fact(DisplayName = "Test of the AddRun method with empty text - should throw FormatException")]
        public void AddRunWithEmptyTextFailingTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            Assert.Throws<FormatException>(() => builder.AddRun(""));
        }

        [Fact(DisplayName = "Test of the AddRun method chaining")]
        public void AddRunChainingTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            FormattedTextBuilder result = builder.AddRun("Hello").AddRun(" ").AddRun("World");
            Assert.Same(builder, result);
            FormattedText formattedText = builder.Build();
            Assert.Equal(3, formattedText.Runs.Count);
            Assert.Equal("Hello World", formattedText.PlainText);
        }

        [Fact(DisplayName = "Test of the AddRun method with style builder")]
        public void AddRunWithStyleBuilderTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            FormattedTextBuilder result = builder.AddRun("Test", sb => sb.Bold());
            Assert.Same(builder, result);
            FormattedText formattedText = builder.Build();
            Assert.Single(formattedText.Runs);
            Assert.NotNull(formattedText.Runs[0].FontStyle);
        }

        [Fact(DisplayName = "Test of the AddRun method with null style builder action")]
        public void AddRunWithNullStyleBuilderActionTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            FormattedTextBuilder result = builder.AddRun("Test", (Action<InlineStyleBuilder>)null);
            Assert.Same(builder, result);
            FormattedText formattedText = builder.Build();
            Assert.Single(formattedText.Runs);
        }

        [Fact(DisplayName = "Test of the AddRun method with style builder and null text - should throw FormatException")]
        public void AddRunWithStyleBuilderAndNullTextFailingTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            Assert.Throws<FormatException>(() => builder.AddRun(null, sb => sb.Bold()));
        }

        [Fact(DisplayName = "Test of the AddRun method with style builder and empty text - should throw FormatException")]
        public void AddRunWithStyleBuilderAndEmptyTextFailingTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            Assert.Throws<FormatException>(() => builder.AddRun("", sb => sb.Bold()));
        }

        [Fact(DisplayName = "Test of the AddRun method with style builder chaining")]
        public void AddRunWithStyleBuilderChainingTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            FormattedTextBuilder result = builder
                .AddRun("Bold", sb => sb.Bold())
                .AddRun(" ")
                .AddRun("Italic", sb => sb.Italic());
            Assert.Same(builder, result);
            FormattedText formattedText = builder.Build();
            Assert.Equal(3, formattedText.Runs.Count);
        }

        [Theory(DisplayName = "Test of the AddPhoneticRun method")]
        [InlineData("ふりがな", 0u, 2u)]
        [InlineData("ピンイン", 5u, 10u)]
        [InlineData("　　", 1u, 3u)]
        [InlineData("0123", 4u, 7u)]
        [InlineData("\t", 100u, 200u)]
        [InlineData("\n\n", 0u, 1u)]
        public void AddPhoneticRunTest(string text, uint startBase, uint endBase)
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            FormattedTextBuilder result = builder.AddPhoneticRun(text, startBase, endBase);
            Assert.Same(builder, result);
            FormattedText formattedText = builder.Build();
            Assert.Single(formattedText.PhoneticRuns);
            Assert.Equal(text, formattedText.PhoneticRuns[0].Text);
            Assert.Equal(startBase, formattedText.PhoneticRuns[0].StartBase);
            Assert.Equal(endBase, formattedText.PhoneticRuns[0].EndBase);
        }

        [Fact(DisplayName = "Test of the AddPhoneticRun method with null text - should throw FormatException")]
        public void AddPhoneticRunWithNullTextFailingTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            Assert.Throws<FormatException>(() => builder.AddPhoneticRun(null, 0, 5));
        }

        [Fact(DisplayName = "Test of the AddPhoneticRun method with empty text - should throw FormatException")]
        public void AddPhoneticRunWithEmptyTextFailingTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            Assert.Throws<FormatException>(() => builder.AddPhoneticRun("", 0, 5));
        }

        [Fact(DisplayName = "Test of the AddPhoneticRun method chaining")]
        public void AddPhoneticRunChainingTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            FormattedTextBuilder result = builder
                .AddPhoneticRun("テスト", 0, 2)
                .AddPhoneticRun("ピンイン", 3, 5);
            Assert.Same(builder, result);
            FormattedText formattedText = builder.Build();
            Assert.Equal(2, formattedText.PhoneticRuns.Count);
        }

        [Theory(DisplayName = "Test of the SetPhoneticProperties method")]
        [InlineData(PhoneticRun.PhoneticType.FullwidthKatakana, PhoneticRun.PhoneticAlignment.Left)]
        [InlineData(PhoneticRun.PhoneticType.HalfwidthKatakana, PhoneticRun.PhoneticAlignment.Center)]
        [InlineData(PhoneticRun.PhoneticType.Hiragana, PhoneticRun.PhoneticAlignment.Distributed)]
        [InlineData(PhoneticRun.PhoneticType.NoConversion, PhoneticRun.PhoneticAlignment.NoControl)]
        public void SetPhoneticPropertiesTest(PhoneticRun.PhoneticType type, PhoneticRun.PhoneticAlignment alignment)
        {
            Font font = new Font();
            FormattedTextBuilder builder = new FormattedTextBuilder();
            FormattedTextBuilder result = builder.SetPhoneticProperties(font, type, alignment);
            Assert.Same(builder, result);
            FormattedText formattedText = builder.Build();
            Assert.NotNull(formattedText.PhoneticProperties);
            Assert.Equal(type, formattedText.PhoneticProperties.Type);
            Assert.Equal(alignment, formattedText.PhoneticProperties.Alignment);
        }

        [Fact(DisplayName = "Test of the SetPhoneticProperties method with default parameters")]
        public void SetPhoneticPropertiesWithDefaultParametersTest()
        {
            Font font = new Font();
            FormattedTextBuilder builder = new FormattedTextBuilder();
            FormattedTextBuilder result = builder.SetPhoneticProperties(font);
            Assert.Same(builder, result);
            FormattedText formattedText = builder.Build();
            Assert.NotNull(formattedText.PhoneticProperties);
            Assert.Equal(PhoneticRun.PhoneticType.FullwidthKatakana, formattedText.PhoneticProperties.Type);
            Assert.Equal(PhoneticRun.PhoneticAlignment.Left, formattedText.PhoneticProperties.Alignment);
        }

        [Fact(DisplayName = "Test of the SetPhoneticProperties method chaining")]
        public void SetPhoneticPropertiesChainingTest()
        {
            Font font = new Font();
            FormattedTextBuilder builder = new FormattedTextBuilder();
            FormattedTextBuilder result = builder
                .AddRun("Test")
                .SetPhoneticProperties(font)
                .AddPhoneticRun("テスト", 0, 2);
            Assert.Same(builder, result);
            FormattedText formattedText = builder.Build();
            Assert.NotNull(formattedText.PhoneticProperties);
        }

        [Fact(DisplayName = "Test of the Build method")]
        public void BuildTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            builder.AddRun("Test");
            FormattedText formattedText = builder.Build();
            Assert.NotNull(formattedText);
            Assert.Single(formattedText.Runs);
            Assert.Equal("Test", formattedText.PlainText);
        }

        [Fact(DisplayName = "Test of the Build method on empty builder")]
        public void BuildOnEmptyBuilderTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            FormattedText formattedText = builder.Build();
            Assert.NotNull(formattedText);
            Assert.Empty(formattedText.Runs);
            Assert.Empty(formattedText.PhoneticRuns);
        }

        [Fact(DisplayName = "Test of the Build method returns same instance")]
        public void BuildReturnsSameInstanceTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            builder.AddRun("Test");
            FormattedText formattedText1 = builder.Build();
            FormattedText formattedText2 = builder.Build();
            Assert.Same(formattedText1, formattedText2);
        }

        [Fact(DisplayName = "Test of the Build method with complex formatted text")]
        public void BuildWithComplexFormattedTextTest()
        {
            Font font1 = new Font();
            font1.Bold = true;
            Font font2 = new Font();
            font2.Italic = true;
            Font phoneticFont = new Font();
            FormattedTextBuilder builder = new FormattedTextBuilder();
            builder
                .AddRun("Bold", font1)
                .AddRun(" ")
                .AddRun("Italic", font2)
                .AddPhoneticRun("テスト", 0, 4)
                .SetPhoneticProperties(phoneticFont, PhoneticRun.PhoneticType.Hiragana);
            FormattedText formattedText = builder.Build();
            Assert.Equal(3, formattedText.Runs.Count);
            Assert.Single(formattedText.PhoneticRuns);
            Assert.NotNull(formattedText.PhoneticProperties);
            Assert.Equal("Bold Italic", formattedText.PlainText);
        }

        [Fact(DisplayName = "Test of the implicit conversion operator")]
        public void ImplicitConversionOperatorTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            builder.AddRun("Test");
            FormattedText formattedText = builder;
            Assert.NotNull(formattedText);
            Assert.Single(formattedText.Runs);
            Assert.Equal("Test", formattedText.PlainText);
        }

        [Fact(DisplayName = "Test of the implicit conversion operator on empty builder")]
        public void ImplicitConversionOperatorOnEmptyBuilderTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            FormattedText formattedText = builder;
            Assert.NotNull(formattedText);
            Assert.Empty(formattedText.Runs);
            Assert.Equal(string.Empty, formattedText.PlainText);
        }

        [Fact(DisplayName = "Test of the implicit conversion operator with complex formatted text")]
        public void ImplicitConversionOperatorWithComplexFormattedTextTest()
        {
            Font font = new Font();
            font.Bold = true;
            FormattedTextBuilder builder = new FormattedTextBuilder();
            builder.AddRun("Hello", font).AddRun(" ").AddRun("World");
            FormattedText formattedText = builder;
            Assert.Equal(3, formattedText.Runs.Count);
            Assert.Equal("Hello World", formattedText.PlainText);
        }

        [Fact(DisplayName = "Test of complete fluent API workflow")]
        public void CompleteFluentApiWorkflowTest()
        {
            Font boldFont = new Font();
            boldFont.Bold = true;
            Font italicFont = new Font();
            italicFont.Italic = true;
            Font phoneticFont = new Font();
            phoneticFont.Size = 8;
            FormattedText formattedText = new FormattedTextBuilder()
                .AddRun("Bold text", boldFont)
                .AddRun(" ")
                .AddRun("Italic text", italicFont)
                .AddPhoneticRun("テスト", 0, 9)
                .SetPhoneticProperties(phoneticFont, PhoneticRun.PhoneticType.Hiragana, PhoneticRun.PhoneticAlignment.Center)
                .Build();
            Assert.NotNull(formattedText);
            Assert.Equal(3, formattedText.Runs.Count);
            Assert.Single(formattedText.PhoneticRuns);
            Assert.NotNull(formattedText.PhoneticProperties);
            Assert.Equal("Bold text Italic text", formattedText.PlainText);
            Assert.Equal(PhoneticRun.PhoneticType.Hiragana, formattedText.PhoneticProperties.Type);
            Assert.Equal(PhoneticRun.PhoneticAlignment.Center, formattedText.PhoneticProperties.Alignment);
        }

        [Fact(DisplayName = "Test of builder modifying underlying FormattedText")]
        public void BuilderModifiesUnderlyingFormattedTextTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            builder.AddRun("First");
            FormattedText formattedText = builder.Build();
            builder.AddRun(" Second");
            Assert.Equal(2, formattedText.Runs.Count);
            Assert.Equal("First Second", formattedText.PlainText);
        }

        [Fact(DisplayName = "Test of multiple Build calls return same instance")]
        public void MultipleBuildCallsReturnSameInstanceTest()
        {
            FormattedTextBuilder builder = new FormattedTextBuilder();
            builder.AddRun("Test");
            FormattedText formattedText1 = builder.Build();
            builder.AddRun(" More");
            FormattedText formattedText2 = builder.Build();
            Assert.Same(formattedText1, formattedText2);
            Assert.Equal("Test More", formattedText2.PlainText);
        }

        [Fact(DisplayName = "Test of mixed method chaining")]
        public void MixedMethodChainingTest()
        {
            Font font = new Font();
            FormattedTextBuilder builder = new FormattedTextBuilder();
            FormattedTextBuilder result = builder
                .AddRun("Run1")
                .AddPhoneticRun("テスト1", 0, 4)
                .AddRun("Run2")
                .SetPhoneticProperties(font)
                .AddPhoneticRun("テスト2", 5, 9)
                .AddRun("Run3");
            Assert.Same(builder, result);
            FormattedText formattedText = builder.Build();
            Assert.Equal(3, formattedText.Runs.Count);
            Assert.Equal(2, formattedText.PhoneticRuns.Count);
            Assert.NotNull(formattedText.PhoneticProperties);
        }

        [Fact(DisplayName = "Test of builder with only phonetic content")]
        public void BuilderWithOnlyPhoneticContentTest()
        {
            Font font = new Font();
            FormattedTextBuilder builder = new FormattedTextBuilder();
            builder
                .AddPhoneticRun("テスト", 0, 2)
                .SetPhoneticProperties(font);
            FormattedText formattedText = builder.Build();
            Assert.Empty(formattedText.Runs);
            Assert.Single(formattedText.PhoneticRuns);
            Assert.NotNull(formattedText.PhoneticProperties);
        }
    }
}
