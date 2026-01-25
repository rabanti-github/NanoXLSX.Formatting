using NanoXLSX.Colors;
using NanoXLSX.Exceptions;
using NanoXLSX.Extensions;
using NanoXLSX.Styles;
using NanoXLSX.Themes;
using Xunit;

namespace NanoXLSX.Formatting.Test.Builder
{
    public class InlineStyleBuilderTest
    {
        [Fact(DisplayName = "Test of the default constructor")]
        public void ConstructorTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            Assert.NotNull(builder);
        }

        [Theory(DisplayName = "Test of the FontName method")]
        [InlineData("Arial")]
        [InlineData("Times New Roman")]
        [InlineData("Calibri")]
        [InlineData("not-existing-font")] // does not validate the font name
        public void FontNameTest(string fontName)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.FontName(fontName);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal(fontName, font.Name);
        }

        [Theory(DisplayName = "Test of the FontName method with null or empty - should throw StyleException")]
        [InlineData("")]
        [InlineData(null)]
        public void FontNameWithNullFailingTest(string fontName)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            Assert.Throws<StyleException>(() => builder.FontName(fontName));
        }

        [Theory(DisplayName = "Test of the Size method")]
        [InlineData(8f)]
        [InlineData(10f)]
        [InlineData(12f)]
        [InlineData(14.5f)]
        public void SizeTest(float size)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Size(size);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal(size, font.Size);
        }

        [Theory(DisplayName = "Test of the Size method with 0 or smaller as value - should fix to min font value automatically")]
        [InlineData(0f)]
        [InlineData(-1f)]
        [InlineData(-0.001)]
        [InlineData(-12f)]
        public void SizeTestWithInvalidSmallTest(int size)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Size(size);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal(Font.MinFontSize, font.Size);
        }

        [Theory(DisplayName = "Test of the Size method with 0 or smaller as value - should fix to max font value automatically")]
        [InlineData(409.1f)]
        [InlineData(500f)]
        [InlineData(800f)]
        public void SizeTestWithInvalidBigTest(int size)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Size(size);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal(Font.MaxFontSize, font.Size);
        }

        [Theory(DisplayName = "Test of the Bold method")]
        [InlineData(true)]
        [InlineData(false)]
        public void BoldTest(bool bold)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Bold(bold);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal(bold, font.Bold);
        }

        [Fact(DisplayName = "Test of the Bold method with default parameter")]
        public void BoldWithDefaultParameterTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Bold();
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.True(font.Bold);
        }

        [Theory(DisplayName = "Test of the Italic method")]
        [InlineData(true)]
        [InlineData(false)]
        public void ItalicTest(bool italic)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Italic(italic);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal(italic, font.Italic);
        }

        [Fact(DisplayName = "Test of the Italic method with default parameter")]
        public void ItalicWithDefaultParameterTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Italic();
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.True(font.Italic);
        }

        [Theory(DisplayName = "Test of the Strikethrough method")]
        [InlineData(true)]
        [InlineData(false)]
        public void StrikethroughTest(bool strikethrough)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Strikethrough(strikethrough);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal(strikethrough, font.Strike);
        }

        [Fact(DisplayName = "Test of the Strikethrough method with default parameter")]
        public void StrikethroughWithDefaultParameterTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Strikethrough();
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.True(font.Strike);
        }

        [Theory(DisplayName = "Test of the Underline method")]
        [InlineData(Font.UnderlineValue.Single)]
        [InlineData(Font.UnderlineValue.Double)]
        [InlineData(Font.UnderlineValue.SingleAccounting)]
        [InlineData(Font.UnderlineValue.DoubleAccounting)]
        [InlineData(Font.UnderlineValue.None)]
        public void UnderlineTest(Font.UnderlineValue underline)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Underline(underline);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal(underline, font.Underline);
        }

        [Fact(DisplayName = "Test of the Underline method with default parameter")]
        public void UnderlineWithDefaultParameterTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Underline();
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal(Font.UnderlineValue.Single, font.Underline);
        }

        [Fact(DisplayName = "Test of the Color method")]
        public void ColorTest()
        {
            Color color = Color.CreateRgb("FF0000");
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Color(color);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal(color, font.ColorValue);
        }

        [Fact(DisplayName = "Test of the Color method with null")]
        public void ColorWithNullTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Color(null);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Null(font.ColorValue);
        }

        [Theory(DisplayName = "Test of the ColorArgb method")]
        [InlineData("FF0000")]
        [InlineData("00FF00")]
        [InlineData("FFFF0000")]
        [InlineData("00FFFF00")]
        public void ColorArgbTest(string argb)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.ColorArgb(argb);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.NotNull(font.ColorValue);
        }

        [Fact(DisplayName = "Test of the ColorTheme method")]
        public void ColorThemeTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.ColorTheme(Theme.ColorSchemeElement.Accent1);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.NotNull(font.ColorValue);
        }

        [Theory(DisplayName = "Test of the ColorTheme method with tint")]
        [InlineData(Theme.ColorSchemeElement.Accent1, 0.0)]
        [InlineData(Theme.ColorSchemeElement.Accent2, 0.5)]
        [InlineData(Theme.ColorSchemeElement.Accent3, -0.5)]
        [InlineData(Theme.ColorSchemeElement.Dark1, 1.0)]
        [InlineData(Theme.ColorSchemeElement.Light1, -1.0)]
        public void ColorThemeWithTintTest(Theme.ColorSchemeElement theme, double tint)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.ColorTheme(theme, tint);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.NotNull(font.ColorValue);
        }

        [Theory(DisplayName = "Test of the ColorIndexed method with enum")]
        [InlineData(IndexedColor.Value.Black)]
        [InlineData(IndexedColor.Value.White)]
        [InlineData(IndexedColor.Value.Red)]
        [InlineData(IndexedColor.Value.Blue)]
        public void ColorIndexedWithEnumTest(IndexedColor.Value indexed)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.ColorIndexed(indexed);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.NotNull(font.ColorValue);
        }

        [Theory(DisplayName = "Test of the ColorIndexed method with int")]
        [InlineData(0)]
        [InlineData(1)]
        [InlineData(32)]
        [InlineData(65)]
        public void ColorIndexedWithIntTest(int colorIndex)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.ColorIndexed(colorIndex);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.NotNull(font.ColorValue);
        }

        [Theory(DisplayName = "Test of the VerticalAlign method")]
        [InlineData(Font.VerticalTextAlignValue.None)]
        [InlineData(Font.VerticalTextAlignValue.Superscript)]
        [InlineData(Font.VerticalTextAlignValue.Subscript)]
        public void VerticalAlignTest(Font.VerticalTextAlignValue alignment)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.VerticalAlign(alignment);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal(alignment, font.VerticalAlign);
        }

        [Fact(DisplayName = "Test of the Superscript method")]
        public void SuperscriptTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Superscript();
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal(Font.VerticalTextAlignValue.Superscript, font.VerticalAlign);
        }

        [Fact(DisplayName = "Test of the Subscript method")]
        public void SubscriptTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Subscript();
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal(Font.VerticalTextAlignValue.Subscript, font.VerticalAlign);
        }

        [Theory(DisplayName = "Test of the Outline method")]
        [InlineData(true)]
        [InlineData(false)]
        public void OutlineTest(bool outline)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Outline(outline);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal(outline, font.Outline);
        }

        [Fact(DisplayName = "Test of the Outline method with default parameter")]
        public void OutlineWithDefaultParameterTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Outline();
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.True(font.Outline);
        }

        [Theory(DisplayName = "Test of the Shadow method")]
        [InlineData(true)]
        [InlineData(false)]
        public void ShadowTest(bool shadow)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Shadow(shadow);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal(shadow, font.Shadow);
        }

        [Fact(DisplayName = "Test of the Shadow method with default parameter")]
        public void ShadowWithDefaultParameterTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Shadow();
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.True(font.Shadow);
        }

        [Theory(DisplayName = "Test of the Condense method")]
        [InlineData(true)]
        [InlineData(false)]
        public void CondenseTest(bool condense)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Condense(condense);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal(condense, font.Condense);
        }

        [Fact(DisplayName = "Test of the Condense method with default parameter")]
        public void CondenseWithDefaultParameterTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Condense();
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.True(font.Condense);
        }

        [Theory(DisplayName = "Test of the Extend method")]
        [InlineData(true)]
        [InlineData(false)]
        public void ExtendTest(bool extend)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Extend(extend);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal(extend, font.Extend);
        }

        [Fact(DisplayName = "Test of the Extend method with default parameter")]
        public void ExtendWithDefaultParameterTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Extend();
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.True(font.Extend);
        }

        [Theory(DisplayName = "Test of the Charset method")]
        [InlineData(Font.CharsetValue.ANSI)]
        [InlineData(Font.CharsetValue.Default)]
        [InlineData(Font.CharsetValue.Symbols)]
        [InlineData(Font.CharsetValue.JIS)]
        public void CharsetTest(Font.CharsetValue charset)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Charset(charset);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal(charset, font.Charset);
        }

        [Theory(DisplayName = "Test of the Family method")]
        [InlineData(Font.FontFamilyValue.NotApplicable)]
        [InlineData(Font.FontFamilyValue.Roman)]
        [InlineData(Font.FontFamilyValue.Swiss)]
        [InlineData(Font.FontFamilyValue.Modern)]
        public void FamilyTest(Font.FontFamilyValue family)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Family(family);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal(family, font.Family);
        }

        [Theory(DisplayName = "Test of the Scheme method")]
        [InlineData(Font.SchemeValue.None)]
        [InlineData(Font.SchemeValue.Major)]
        [InlineData(Font.SchemeValue.Minor)]
        public void SchemeTest(Font.SchemeValue scheme)
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder.Scheme(scheme);
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal(scheme, font.Scheme);
        }

        [Fact(DisplayName = "Test of the Build method")]
        public void BuildTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            builder.Bold().Italic();
            Font font = builder.Build();
            Assert.NotNull(font);
            Assert.True(font.Bold);
            Assert.True(font.Italic);
        }

        [Fact(DisplayName = "Test of the Build method with reset false")]
        public void BuildWithResetFalseTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            builder.Bold().Italic();
            Font font1 = builder.Build(false);
            Font font2 = builder.Build(false);
            Assert.NotSame(font1, font2);
            Assert.True(font1.Bold);
            Assert.True(font1.Italic);
            Assert.True(font2.Bold);
            Assert.True(font2.Italic);
        }

        [Fact(DisplayName = "Test of the Build method with reset true")]
        public void BuildWithResetTrueTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            builder.Bold().Italic();
            Font font1 = builder.Build(true);
            Font font2 = builder.Build(true);
            Assert.NotSame(font1, font2);
            Assert.True(font1.Bold);
            Assert.True(font1.Italic);
            Assert.False(font2.Bold);
            Assert.False(font2.Italic);
        }

        [Fact(DisplayName = "Test of the Build method default parameter resets builder")]
        public void BuildDefaultParameterResetsBuilderTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            builder.Bold();
            Font font1 = builder.Build();
            builder.Italic();
            Font font2 = builder.Build();
            Assert.True(font1.Bold);
            Assert.False(font1.Italic);
            Assert.False(font2.Bold);
            Assert.True(font2.Italic);
        }

        [Fact(DisplayName = "Test of the Build method creates independent copy")]
        public void BuildCreatesIndependentCopyTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            builder.Bold();
            Font font = builder.Build(false);
            font.Italic = true;
            Font font2 = builder.Build(false);
            Assert.True(font.Italic);
            Assert.False(font2.Italic);
        }

        [Fact(DisplayName = "Test of method chaining")]
        public void MethodChainingTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            InlineStyleBuilder result = builder
                .FontName("Arial")
                .Size(12)
                .Bold()
                .Italic()
                .Underline();
            Assert.Same(builder, result);
            Font font = builder.Build(false);
            Assert.Equal("Arial", font.Name);
            Assert.Equal(12, font.Size);
            Assert.True(font.Bold);
            Assert.True(font.Italic);
            Assert.Equal(Font.UnderlineValue.Single, font.Underline);
        }

        [Fact(DisplayName = "Test of complete fluent API workflow")]
        public void CompleteFluentApiWorkflowTest()
        {
            Color red = Color.CreateRgb("FF0000");
            Font font = new InlineStyleBuilder()
                .FontName("Calibri")
                .Size(14)
                .Bold()
                .Italic()
                .Strikethrough()
                .Underline(Font.UnderlineValue.Double)
                .Color(red)
                .Superscript()
                .Outline()
                .Shadow()
                .Build();
            Assert.Equal("Calibri", font.Name);
            Assert.Equal(14, font.Size);
            Assert.True(font.Bold);
            Assert.True(font.Italic);
            Assert.True(font.Strike);
            Assert.Equal(Font.UnderlineValue.Double, font.Underline);
            Assert.Equal(red, font.ColorValue);
            Assert.Equal(Font.VerticalTextAlignValue.Superscript, font.VerticalAlign);
            Assert.True(font.Outline);
            Assert.True(font.Shadow);
        }

        [Fact(DisplayName = "Test of builder reusability after reset")]
        public void BuilderReusabilityAfterResetTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            Font font1 = builder.Bold().Build();
            Font font2 = builder.Italic().Build();
            Font font3 = builder.Underline().Build();
            Assert.True(font1.Bold);
            Assert.False(font1.Italic);
            Assert.False(font2.Bold);
            Assert.True(font2.Italic);
            Assert.False(font3.Bold);
            Assert.False(font3.Italic);
            Assert.Equal(Font.UnderlineValue.Single, font3.Underline);
        }

        [Fact(DisplayName = "Test of empty builder Build")]
        public void EmptyBuilderBuildTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            Font font = builder.Build();
            Assert.NotNull(font);
            Assert.False(font.Bold);
            Assert.False(font.Italic);
        }

        [Fact(DisplayName = "Test of overwriting properties")]
        public void OverwritingPropertiesTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            builder.Bold(true).Bold(false);
            builder.FontName("Arial").FontName("Calibri");
            builder.Size(10).Size(12);
            Font font = builder.Build(false);
            Assert.False(font.Bold);
            Assert.Equal("Calibri", font.Name);
            Assert.Equal(12, font.Size);
        }

        [Fact(DisplayName = "Test of all color methods")]
        public void AllColorMethodsTest()
        {
            Color directColor = Color.CreateRgb("FF0000");
            InlineStyleBuilder builder1 = new InlineStyleBuilder();
            Font font1 = builder1.Color(directColor).Build();
            Assert.NotNull(font1.ColorValue);
            InlineStyleBuilder builder2 = new InlineStyleBuilder();
            Font font2 = builder2.ColorArgb("00FF00").Build();
            Assert.NotNull(font2.ColorValue);
            InlineStyleBuilder builder3 = new InlineStyleBuilder();
            Font font3 = builder3.ColorTheme(Theme.ColorSchemeElement.Accent1, 0.5).Build();
            Assert.NotNull(font3.ColorValue);
            InlineStyleBuilder builder4 = new InlineStyleBuilder();
            Font font4 = builder4.ColorIndexed(IndexedColor.Value.Red).Build();
            Assert.NotNull(font4.ColorValue);
            InlineStyleBuilder builder5 = new InlineStyleBuilder();
            Font font5 = builder5.ColorIndexed(10).Build();
            Assert.NotNull(font5.ColorValue);
        }

        [Fact(DisplayName = "Test of all vertical alignment methods")]
        public void AllVerticalAlignmentMethodsTest()
        {
            InlineStyleBuilder builder1 = new InlineStyleBuilder();
            Font font1 = builder1.Superscript().Build();
            Assert.Equal(Font.VerticalTextAlignValue.Superscript, font1.VerticalAlign);
            InlineStyleBuilder builder2 = new InlineStyleBuilder();
            Font font2 = builder2.Subscript().Build();
            Assert.Equal(Font.VerticalTextAlignValue.Subscript, font2.VerticalAlign);
            InlineStyleBuilder builder3 = new InlineStyleBuilder();
            Font font3 = builder3.VerticalAlign(Font.VerticalTextAlignValue.None).Build();
            Assert.Equal(Font.VerticalTextAlignValue.None, font3.VerticalAlign);
        }

        [Fact(DisplayName = "Test of mixing boolean property states")]
        public void MixingBooleanPropertyStatesTest()
        {
            InlineStyleBuilder builder = new InlineStyleBuilder();
            Font font = builder
                .Bold(true)
                .Italic(false)
                .Strikethrough(true)
                .Outline(false)
                .Shadow(true)
                .Condense(false)
                .Extend(true)
                .Build(false);
            Assert.True(font.Bold);
            Assert.False(font.Italic);
            Assert.True(font.Strike);
            Assert.False(font.Outline);
            Assert.True(font.Shadow);
            Assert.False(font.Condense);
            Assert.True(font.Extend);
        }

        [Fact(DisplayName = "Test of complex chaining with multiple property types")]
        public void ComplexChainingWithMultiplePropertyTypesTest()
        {
            Font font = new InlineStyleBuilder()
                .FontName("Times New Roman")
                .Size(16)
                .Bold()
                .ColorArgb("0000FF")
                .Underline(Font.UnderlineValue.DoubleAccounting)
                .VerticalAlign(Font.VerticalTextAlignValue.Subscript)
                .Charset(Font.CharsetValue.ANSI)
                .Family(Font.FontFamilyValue.Roman)
                .Scheme(Font.SchemeValue.Major)
                .Build();
            Assert.Equal("Times New Roman", font.Name);
            Assert.Equal(16, font.Size);
            Assert.True(font.Bold);
            Assert.NotNull(font.ColorValue);
            Assert.Equal(Font.UnderlineValue.DoubleAccounting, font.Underline);
            Assert.Equal(Font.VerticalTextAlignValue.Subscript, font.VerticalAlign);
            Assert.Equal(Font.CharsetValue.ANSI, font.Charset);
            Assert.Equal(Font.FontFamilyValue.Roman, font.Family);
            Assert.Equal(Font.SchemeValue.Major, font.Scheme);
        }
    }
}
