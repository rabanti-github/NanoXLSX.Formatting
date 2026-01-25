namespace NanoXLSX.Formatting.Test.Data
{
    using global::NanoXLSX.Extensions;
    using global::NanoXLSX.Styles;
    using Xunit;
    using FormatException = Exceptions.FormatException;

    namespace NanoXLSX.Tests
    {
        public class TextRunTests
        {
            [Fact(DisplayName = "Test of the constructor with text only")]
            public void ConstructorWithTextOnlyTest()
            {
                TextRun run = new TextRun("Test");

                Assert.NotNull(run);
                Assert.Equal("Test", run.Text);
                Assert.Null(run.FontStyle);
            }

            [Fact(DisplayName = "Test of the constructor with text and font style")]
            public void ConstructorWithTextAndFontTest()
            {
                Font font = new Font();
                font.Bold = true;
                TextRun run = new TextRun("Test", font);

                Assert.NotNull(run);
                Assert.Equal("Test", run.Text);
                Assert.NotNull(run.FontStyle);
                Assert.Same(font, run.FontStyle);
            }

            [Fact(DisplayName = "Test of the constructor with text and null font style")]
            public void ConstructorWithTextAndNullFontTest()
            {
                TextRun run = new TextRun("Test", null);

                Assert.NotNull(run);
                Assert.Equal("Test", run.Text);
                Assert.Null(run.FontStyle);
            }

            [Theory(DisplayName = "Test of the constructor with various text values")]
            [InlineData("Hello")]
            [InlineData("World")]
            [InlineData("123")]
            [InlineData(" ")]
            [InlineData("Special @#$%")]
            public void ConstructorWithVariousTextValuesTest(string text)
            {
                TextRun run = new TextRun(text);

                Assert.Equal(text, run.Text);
            }

            [Fact(DisplayName = "Test of the constructor with null text - should throw FormatException")]
            public void ConstructorWithNullTextFailingTest()
            {
                Assert.Throws<FormatException>(() => new TextRun(null));
            }

            [Fact(DisplayName = "Test of the constructor with null text and font - should throw FormatException")]
            public void ConstructorWithNullTextAndFontFailingTest()
            {
                Font font = new Font();
                Assert.Throws<FormatException>(() => new TextRun(null, font));
            }

            [Fact(DisplayName = "Test of the Text property getter")]
            public void TextPropertyGetterTest()
            {
                TextRun run = new TextRun("Initial");
                Assert.Equal("Initial", run.Text);
            }

            [Theory(DisplayName = "Test of the Text property setter")]
            [InlineData("Modified")]
            [InlineData("Another value")]
            [InlineData("")]
            [InlineData("123")]
            [InlineData("テスト")]
            [InlineData("你好")]
            [InlineData("　　")]
            [InlineData("\t")]
            [InlineData("\r\n")]
            public void TextPropertySetterTest(string newText)
            {
                TextRun run = new TextRun("Initial");
                Assert.Equal("Initial", run.Text);
                run.Text = newText;
                Assert.Equal(newText, run.Text);
            }

            [Fact(DisplayName = "Test of the Text property setter with null  - should throw FormatException")]
            public void TextPropertySetterWithNullFailingTest()
            {
                TextRun run = new TextRun("Initial");
                Assert.Throws<FormatException>(() => run.Text = null);
            }

            [Fact(DisplayName = "Test of the FontStyle property getter")]
            public void FontStylePropertyGetterTest()
            {
                Font font = new Font();
                font.Italic = true;
                TextRun run = new TextRun("Test", font);
                Assert.Same(font, run.FontStyle);
            }

            [Fact(DisplayName = "Test of the FontStyle property setter")]
            public void FontStylePropertySetterTest()
            {
                TextRun run = new TextRun("Test");
                Font newFont = new Font();
                newFont.Bold = true;
                run.FontStyle = newFont;
                Assert.Same(newFont, run.FontStyle);
            }

            [Fact(DisplayName = "Test of the FontStyle property setter with null")]
            public void FontStylePropertySetterWithNullTest()
            {
                Font font = new Font();
                TextRun run = new TextRun("Test", font);
                run.FontStyle = null;
                Assert.Null(run.FontStyle);
            }

            [Fact(DisplayName = "Test of the FontStyle property setter replacing existing font")]
            public void FontStylePropertySetterReplacingTest()
            {
                Font font1 = new Font();
                font1.Bold = true;
                Font font2 = new Font();
                font2.Italic = true;
                TextRun run = new TextRun("Test", font1);
                run.FontStyle = font2;
                Assert.Same(font2, run.FontStyle);
                Assert.NotSame(font1, run.FontStyle);
            }

            [Fact(DisplayName = "Test of the Copy method")]
            public void CopyTest()
            {
                Font font = new Font();
                font.Bold = true;
                TextRun original = new TextRun("Original", font);
                TextRun copy = original.Copy();
                Assert.NotSame(original, copy);
                Assert.Equal(original.Text, copy.Text);
            }

            [Fact(DisplayName = "Test of the Copy method with null font style")]
            public void CopyWithNullFontTest()
            {
                TextRun original = new TextRun("Original", null);
                TextRun copy = original.Copy();
                Assert.NotSame(original, copy);
                Assert.Equal(original.Text, copy.Text);
                Assert.Null(copy.FontStyle);
            }

            [Fact(DisplayName = "Test of the Copy method creates deep copy")]
            public void CopyCreatesDeepCopyTest()
            {
                Font font = new Font();
                font.Bold = true;
                TextRun original = new TextRun("Original", font);
                TextRun copy = original.Copy();
                copy.Text = "Modified";
                Assert.Equal("Original", original.Text);
                Assert.Equal("Modified", copy.Text);
            }

            [Fact(DisplayName = "Test of the Copy method creates independent font copy")]
            public void CopyCreatesIndependentFontCopyTest()
            {
                Font font = new Font();
                font.Bold = true;
                TextRun original = new TextRun("Test", font);
                TextRun copy = original.Copy();
                Assert.NotNull(copy.FontStyle);
                Assert.NotSame(original.FontStyle, copy.FontStyle);
            }

            [Theory(DisplayName = "Test of the Copy method with various text values")]
            [InlineData("Test")]
            [InlineData("Hello World")]
            [InlineData(" ")]
            [InlineData("Special chars: @#$%")]
            [InlineData("")]
            [InlineData("123")]
            [InlineData("テスト")]
            [InlineData("你好")]
            [InlineData("\t")]
            [InlineData("\r\n")]
            public void CopyWithVariousTextValuesTest(string text)
            {
                TextRun original = new TextRun(text);
                TextRun copy = original.Copy();
                Assert.Equal(text, copy.Text);
            }

            [Fact(DisplayName = "Test of the Copy method font style independence")]
            public void CopyFontStyleIndependenceTest()
            {
                Font font = new Font();
                font.Bold = true;
                TextRun original = new TextRun("Test", font);
                TextRun copy = original.Copy();
                copy.FontStyle.Bold = false;
                Assert.True(original.FontStyle.Bold);
                Assert.False(copy.FontStyle.Bold);
            }
        }
    }
}
