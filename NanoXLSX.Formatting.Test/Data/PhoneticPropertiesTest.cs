using NanoXLSX.Extensions;
using NanoXLSX.Styles;
using Xunit;

namespace NanoXLSX.Formatting.Test.Data
{
    public class PhoneticPropertiesTest
    {
        [Fact(DisplayName = "Test of the constructor with font reference")]
        public void ConstructorWithFontReferenceTest()
        {
            Font font = new Font();
            PhoneticProperties properties = new PhoneticProperties(font);
            Assert.NotNull(properties);
            Assert.Same(font, properties.FontReference);
            Assert.Equal(PhoneticRun.PhoneticType.FullwidthKatakana, properties.Type);
            Assert.Equal(PhoneticRun.PhoneticAlignment.Left, properties.Alignment);
            Assert.Equal(0, properties.FontId);
        }

        [Fact(DisplayName = "Test of the constructor with null font reference")]
        public void ConstructorWithNullFontReferenceTest()
        {
            PhoneticProperties properties = new PhoneticProperties(null);
            Assert.NotNull(properties);
            Assert.Null(properties.FontReference);
            Assert.Equal(PhoneticRun.PhoneticType.FullwidthKatakana, properties.Type);
            Assert.Equal(PhoneticRun.PhoneticAlignment.Left, properties.Alignment);
        }

        [Fact(DisplayName = "Test of the constructor with styled font")]
        public void ConstructorWithStyledFontTest()
        {
            Font font = new Font();
            font.Bold = true;
            font.Italic = true;
            PhoneticProperties properties = new PhoneticProperties(font);
            Assert.Same(font, properties.FontReference);
            Assert.True(properties.FontReference.Bold);
            Assert.True(properties.FontReference.Italic);
        }

        [Fact(DisplayName = "Test of the FontReference property getter")]
        public void FontReferencePropertyGetterTest()
        {
            Font font = new Font();
            font.Bold = true;
            PhoneticProperties properties = new PhoneticProperties(font);
            Assert.Same(font, properties.FontReference);
        }

        [Fact(DisplayName = "Test of the FontReference property setter")]
        public void FontReferencePropertySetterTest()
        {
            Font font1 = new Font();
            Font font2 = new Font();
            font2.Italic = true;
            PhoneticProperties properties = new PhoneticProperties(font1);
            properties.FontReference = font2;
            Assert.Same(font2, properties.FontReference);
            Assert.NotSame(font1, properties.FontReference);
        }

        [Fact(DisplayName = "Test of the FontReference property setter with null")]
        public void FontReferencePropertySetterWithNullTest()
        {
            Font font = new Font();
            PhoneticProperties properties = new PhoneticProperties(font);
            properties.FontReference = null;
            Assert.Null(properties.FontReference);
        }

        [Fact(DisplayName = "Test of the FontId property getter")]
        public void FontIdPropertyGetterTest()
        {
            Font font = new Font();
            PhoneticProperties properties = new PhoneticProperties(font);
            Assert.Equal(0, properties.FontId);
        }

        [Theory(DisplayName = "Test of the FontId property setter")]
        [InlineData(0)]
        [InlineData(1)]
        [InlineData(5)]
        [InlineData(100)]
        [InlineData(-1)]
        public void FontIdPropertySetterTest(int fontId)
        {
            Font font = new Font();
            PhoneticProperties properties = new PhoneticProperties(font);
            properties.FontId = fontId;
            Assert.Equal(fontId, properties.FontId);
        }

        [Fact(DisplayName = "Test of the Type property default value")]
        public void TypePropertyDefaultValueTest()
        {
            Font font = new Font();
            PhoneticProperties properties = new PhoneticProperties(font);
            Assert.Equal(PhoneticRun.PhoneticType.FullwidthKatakana, properties.Type);
        }

        [Theory(DisplayName = "Test of the Type property setter")]
        [InlineData(PhoneticRun.PhoneticType.HalfwidthKatakana)]
        [InlineData(PhoneticRun.PhoneticType.FullwidthKatakana)]
        [InlineData(PhoneticRun.PhoneticType.Hiragana)]
        [InlineData(PhoneticRun.PhoneticType.NoConversion)]
        public void TypePropertySetterTest(PhoneticRun.PhoneticType type)
        {
            Font font = new Font();
            PhoneticProperties properties = new PhoneticProperties(font);
            properties.Type = type;
            Assert.Equal(type, properties.Type);
        }

        [Fact(DisplayName = "Test of the Type property getter")]
        public void TypePropertyGetterTest()
        {
            Font font = new Font();
            PhoneticProperties properties = new PhoneticProperties(font);
            properties.Type = PhoneticRun.PhoneticType.Hiragana;
            Assert.Equal(PhoneticRun.PhoneticType.Hiragana, properties.Type);
        }

        [Fact(DisplayName = "Test of the Alignment property default value")]
        public void AlignmentPropertyDefaultValueTest()
        {
            Font font = new Font();
            PhoneticProperties properties = new PhoneticProperties(font);
            Assert.Equal(PhoneticRun.PhoneticAlignment.Left, properties.Alignment);
        }

        [Theory(DisplayName = "Test of the Alignment property setter")]
        [InlineData(PhoneticRun.PhoneticAlignment.NoControl)]
        [InlineData(PhoneticRun.PhoneticAlignment.Left)]
        [InlineData(PhoneticRun.PhoneticAlignment.Center)]
        [InlineData(PhoneticRun.PhoneticAlignment.Distributed)]
        public void AlignmentPropertySetterTest(PhoneticRun.PhoneticAlignment alignment)
        {
            Font font = new Font();
            PhoneticProperties properties = new PhoneticProperties(font);
            properties.Alignment = alignment;
            Assert.Equal(alignment, properties.Alignment);
        }

        [Fact(DisplayName = "Test of the Alignment property getter")]
        public void AlignmentPropertyGetterTest()
        {
            Font font = new Font();
            PhoneticProperties properties = new PhoneticProperties(font);
            properties.Alignment = PhoneticRun.PhoneticAlignment.Center;
            Assert.Equal(PhoneticRun.PhoneticAlignment.Center, properties.Alignment);
        }

        [Fact(DisplayName = "Test of the Copy method")]
        public void CopyTest()
        {
            Font font = new Font();
            font.Bold = true;
            PhoneticProperties original = new PhoneticProperties(font);
            original.Type = PhoneticRun.PhoneticType.Hiragana;
            original.Alignment = PhoneticRun.PhoneticAlignment.Center;
            original.FontId = 5;
            PhoneticProperties copy = original.Copy();
            Assert.NotSame(original, copy);
            Assert.Same(original.FontReference, copy.FontReference);
            Assert.Equal(original.Type, copy.Type);
            Assert.Equal(original.Alignment, copy.Alignment);
        }

        [Fact(DisplayName = "Test of the Copy method with null font reference")]
        public void CopyWithNullFontReferenceTest()
        {
            PhoneticProperties original = new PhoneticProperties(null);
            original.Type = PhoneticRun.PhoneticType.HalfwidthKatakana;
            original.Alignment = PhoneticRun.PhoneticAlignment.Distributed;
            PhoneticProperties copy = original.Copy();
            Assert.NotSame(original, copy);
            Assert.Null(copy.FontReference);
            Assert.Equal(original.Type, copy.Type);
            Assert.Equal(original.Alignment, copy.Alignment);
        }

        [Fact(DisplayName = "Test of the Copy method with default values")]
        public void CopyWithDefaultValuesTest()
        {
            Font font = new Font();
            PhoneticProperties original = new PhoneticProperties(font);
            PhoneticProperties copy = original.Copy();
            Assert.NotSame(original, copy);
            Assert.Equal(PhoneticRun.PhoneticType.FullwidthKatakana, copy.Type);
            Assert.Equal(PhoneticRun.PhoneticAlignment.Left, copy.Alignment);
        }

        [Fact(DisplayName = "Test of the Copy method creates independent instance")]
        public void CopyCreatesIndependentInstanceTest()
        {
            Font font = new Font();
            PhoneticProperties original = new PhoneticProperties(font);
            original.Type = PhoneticRun.PhoneticType.Hiragana;
            original.Alignment = PhoneticRun.PhoneticAlignment.Center;
            PhoneticProperties copy = original.Copy();
            copy.Type = PhoneticRun.PhoneticType.NoConversion;
            copy.Alignment = PhoneticRun.PhoneticAlignment.Distributed;
            Assert.Equal(PhoneticRun.PhoneticType.Hiragana, original.Type);
            Assert.Equal(PhoneticRun.PhoneticAlignment.Center, original.Alignment);
            Assert.Equal(PhoneticRun.PhoneticType.NoConversion, copy.Type);
            Assert.Equal(PhoneticRun.PhoneticAlignment.Distributed, copy.Alignment);
        }

        [Fact(DisplayName = "Test of the Copy method does not copy FontId")]
        public void CopyDoesNotCopyFontIdTest()
        {
            Font font = new Font();
            PhoneticProperties original = new PhoneticProperties(font);
            original.FontId = 42;
            PhoneticProperties copy = original.Copy();
            Assert.Equal(0, copy.FontId);
            Assert.Equal(42, original.FontId);
        }

        [Fact(DisplayName = "Test of the Copy method shares font reference")]
        public void CopySharesFontReferenceTest()
        {
            Font font = new Font();
            font.Bold = true;
            PhoneticProperties original = new PhoneticProperties(font);
            PhoneticProperties copy = original.Copy();
            copy.FontReference.Italic = true;
            Assert.Same(original.FontReference, copy.FontReference);
            Assert.True(original.FontReference.Italic);
            Assert.True(copy.FontReference.Italic);
        }

        [Theory(DisplayName = "Test of complete property combination")]
        [InlineData(PhoneticRun.PhoneticType.HalfwidthKatakana, PhoneticRun.PhoneticAlignment.NoControl, 0)]
        [InlineData(PhoneticRun.PhoneticType.FullwidthKatakana, PhoneticRun.PhoneticAlignment.Left, 1)]
        [InlineData(PhoneticRun.PhoneticType.Hiragana, PhoneticRun.PhoneticAlignment.Center, 5)]
        [InlineData(PhoneticRun.PhoneticType.NoConversion, PhoneticRun.PhoneticAlignment.Distributed, 10)]
        public void CompletePropertyCombinationTest(PhoneticRun.PhoneticType type, PhoneticRun.PhoneticAlignment alignment, int fontId)
        {
            Font font = new Font();
            PhoneticProperties properties = new PhoneticProperties(font);
            properties.Type = type;
            properties.Alignment = alignment;
            properties.FontId = fontId;
            Assert.Equal(type, properties.Type);
            Assert.Equal(alignment, properties.Alignment);
            Assert.Equal(fontId, properties.FontId);
            Assert.Same(font, properties.FontReference);
        }

        [Fact(DisplayName = "Test of modifying properties after construction")]
        public void ModifyingPropertiesAfterConstructionTest()
        {
            Font font1 = new Font();
            font1.Bold = true;
            Font font2 = new Font();
            font2.Italic = true;
            PhoneticProperties properties = new PhoneticProperties(font1);
            properties.FontReference = font2;
            properties.Type = PhoneticRun.PhoneticType.Hiragana;
            properties.Alignment = PhoneticRun.PhoneticAlignment.Center;
            properties.FontId = 7;
            Assert.Same(font2, properties.FontReference);
            Assert.Equal(PhoneticRun.PhoneticType.Hiragana, properties.Type);
            Assert.Equal(PhoneticRun.PhoneticAlignment.Center, properties.Alignment);
            Assert.Equal(7, properties.FontId);
        }

        [Fact(DisplayName = "Test of the Equals method for equality")]
        public void EqualsTestForEquality()
        {
            Font font1 = new Font();
            font1.Bold = true;
            font1.Size = 12;
            font1.Name = "Arial";

            Font font2 = new Font();
            font2.Bold = true;
            font2.Size = 12;
            font2.Name = "Arial";

            PhoneticProperties props1 = new PhoneticProperties(font1);
            props1.Type = PhoneticRun.PhoneticType.Hiragana;
            props1.Alignment = PhoneticRun.PhoneticAlignment.Center;

            PhoneticProperties props2 = new PhoneticProperties(font2);
            props2.Type = PhoneticRun.PhoneticType.Hiragana;
            props2.Alignment = PhoneticRun.PhoneticAlignment.Center;

            Assert.True(props1.Equals(props2));
        }

        [Fact(DisplayName = "Test of the Equals method for inequality")]
        public void EqualsTestForInequality()
        {
            Font font1 = new Font();
            font1.Bold = true;
            font1.Size = 12;

            Font font2 = new Font();
            font2.Bold = false;
            font2.Size = 12;

            PhoneticProperties props1 = new PhoneticProperties(font1);
            props1.Type = PhoneticRun.PhoneticType.Hiragana;
            props1.Alignment = PhoneticRun.PhoneticAlignment.Center;

            PhoneticProperties props2 = new PhoneticProperties(font2);
            props2.Type = PhoneticRun.PhoneticType.Hiragana;
            props2.Alignment = PhoneticRun.PhoneticAlignment.Center;

            Assert.False(props1.Equals(props2));
        }

        [Fact(DisplayName = "Test of the Equals method for inequality on different object types")]
        public void EqualsTestForInequality2()
        {
            Font font1 = new Font();
            font1.Bold = true;
            font1.Size = 12;

            PhoneticProperties props1 = new PhoneticProperties(font1);
            props1.Type = PhoneticRun.PhoneticType.Hiragana;
            props1.Alignment = PhoneticRun.PhoneticAlignment.Center;

            string obj2 = "This is a string object";

            Assert.False(props1.Equals(obj2));
        }

        [Fact(DisplayName = "Test of the GetHashCode method for equality")]
        public void GetHashCodeTestForEquality()
        {
            Font font1 = new Font();
            font1.Bold = true;
            font1.Size = 12;
            font1.Name = "Arial";

            Font font2 = new Font();
            font2.Bold = true;
            font2.Size = 12;
            font2.Name = "Arial";

            PhoneticProperties props1 = new PhoneticProperties(font1);
            props1.Type = PhoneticRun.PhoneticType.Hiragana;
            props1.Alignment = PhoneticRun.PhoneticAlignment.Center;

            PhoneticProperties props2 = new PhoneticProperties(font2);
            props2.Type = PhoneticRun.PhoneticType.Hiragana;
            props2.Alignment = PhoneticRun.PhoneticAlignment.Center;

            Assert.Equal(props1.GetHashCode(), props2.GetHashCode());
        }

        [Fact(DisplayName = "Test of the GetHashCode method for inequality")]
        public void GetHashCodeTestForInequality()
        {
            Font font1 = new Font();
            font1.Bold = true;

            Font font2 = new Font();
            font2.Bold = true;

            PhoneticProperties props1 = new PhoneticProperties(font1);
            props1.Type = PhoneticRun.PhoneticType.Hiragana;
            props1.Alignment = PhoneticRun.PhoneticAlignment.Center;

            PhoneticProperties props2 = new PhoneticProperties(font2);
            props2.Type = PhoneticRun.PhoneticType.FullwidthKatakana;
            props2.Alignment = PhoneticRun.PhoneticAlignment.Left;

            Assert.NotEqual(props1.GetHashCode(), props2.GetHashCode());
        }
    }
}
