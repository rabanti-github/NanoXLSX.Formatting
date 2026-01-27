using NanoXLSX.Exceptions;
using NanoXLSX.Extensions;
using Xunit;

namespace NanoXLSX.Formatting.Test.Data
{
    public class PhoneticRunTest
    {
        [Fact(DisplayName = "Test of the constructor")]
        public void ConstructorTest()
        {
            PhoneticRun run = new PhoneticRun("テスト", 0, 5);
            Assert.NotNull(run);
            Assert.Equal("テスト", run.Text);
            Assert.Equal(0u, run.StartBase);
            Assert.Equal(5u, run.EndBase);
        }

        [Theory(DisplayName = "Test of the constructor with various parameters")]
        [InlineData("ふりがな", 0u, 2u)]
        [InlineData("ピンイン", 5u, 10u)]
        [InlineData("かな", 0u, 0u)]
        [InlineData("テスト", 100u, 200u)]
        public void ConstructorWithVariousParametersTest(string text, uint startBase, uint endBase)
        {
            PhoneticRun run = new PhoneticRun(text, startBase, endBase);
            Assert.Equal(text, run.Text);
            Assert.Equal(startBase, run.StartBase);
            Assert.Equal(endBase, run.EndBase);
        }

        [Fact(DisplayName = "Test of the constructor with null text - should throw FormatException")]
        public void ConstructorWithNullTextFailingTest()
        {
            Assert.Throws<FormatException>(() => new PhoneticRun(null, 0, 5));
        }

        [Fact(DisplayName = "Test of the constructor with empty text - should throw FormatException")]
        public void ConstructorWithEmptyTextFailingTest()
        {
            Assert.Throws<FormatException>(() => new PhoneticRun("", 0, 5));
        }

        [Theory(DisplayName = "Test of the constructor with invalid text - should throw FormatException")]
        [InlineData(null)]
        [InlineData("")]
        public void ConstructorWithInvalidTextFailingTest(string text)
        {
            Assert.Throws<FormatException>(() => new PhoneticRun(text, 0, 5));
        }

        [Fact(DisplayName = "Test of the Text property getter")]
        public void TextPropertyGetterTest()
        {
            PhoneticRun run = new PhoneticRun("テスト", 0, 5);
            Assert.Equal("テスト", run.Text);
        }

        [Theory(DisplayName = "Test of the Text property setter")]
        [InlineData("ふりがな")]
        [InlineData("ピンイン")]
        [InlineData("　　")]
        [InlineData("0123")]
        [InlineData("\t")]
        [InlineData("\n\n")]
        public void TextPropertySetterTest(string newText)
        {
            PhoneticRun run = new PhoneticRun("Initial", 0, 5);
            Assert.Equal("Initial", run.Text);
            run.Text = newText;
            Assert.Equal(newText, run.Text);
        }

        [Theory(DisplayName = "Test of the Text property setter with null or empty - should throw FormatException")]
        [InlineData("")]
        [InlineData(null)]
        public void TextPropertySetterWithNullFailingTest(string givebValue)
        {
            PhoneticRun run = new PhoneticRun("Initial", 0, 5);
            Assert.Throws<FormatException>(() => run.Text = givebValue);
        }

        [Fact(DisplayName = "Test of the StartBase property getter")]
        public void StartBasePropertyGetterTest()
        {
            PhoneticRun run = new PhoneticRun("テスト", 10, 20);
            Assert.Equal(10u, run.StartBase);
        }

        [Theory(DisplayName = "Test of the StartBase property setter")]
        [InlineData(0u)]
        [InlineData(5u)]
        [InlineData(100u)]
        [InlineData(uint.MaxValue)]
        public void StartBasePropertySetterTest(uint newStartBase)
        {
            PhoneticRun run = new PhoneticRun("テスト", 1, 5);
            Assert.Equal(1u, run.StartBase);
            run.StartBase = newStartBase;
            Assert.Equal(newStartBase, run.StartBase);
        }

        [Fact(DisplayName = "Test of the EndBase property getter")]
        public void EndBasePropertyGetterTest()
        {
            PhoneticRun run = new PhoneticRun("テスト", 10, 20);
            Assert.Equal(20u, run.EndBase);
        }

        [Theory(DisplayName = "Test of the EndBase property setter")]
        [InlineData(0u)]
        [InlineData(5u)]
        [InlineData(100u)]
        [InlineData(uint.MaxValue)]
        public void EndBasePropertySetterTest(uint newEndBase)
        {
            PhoneticRun run = new PhoneticRun("テスト", 0, 1);
            Assert.Equal(1u, run.EndBase);
            run.EndBase = newEndBase;
            Assert.Equal(newEndBase, run.EndBase);
        }

        [Fact(DisplayName = "Test of the Copy method")]
        public void CopyTest()
        {
            PhoneticRun original = new PhoneticRun("テスト", 5, 10);
            PhoneticRun copy = original.Copy();
            Assert.NotSame(original, copy);
            Assert.Equal(original.Text, copy.Text);
            Assert.Equal(original.StartBase, copy.StartBase);
            Assert.Equal(original.EndBase, copy.EndBase);
        }

        [Theory(DisplayName = "Test of the Copy method with various values")]
        [InlineData("ふりがな", 0u, 2u)]
        [InlineData("ピンイン", 5u, 10u)]
        [InlineData("　　", 1u, 3u)]
        [InlineData("0123", 4u, 7u)]
        [InlineData("\t", 100u, 200u)]
        [InlineData("\n\n", 0u, 1u)]
        public void CopyWithVariousValuesTest(string text, uint startBase, uint endBase)
        {
            PhoneticRun original = new PhoneticRun(text, startBase, endBase);
            PhoneticRun copy = original.Copy();
            Assert.NotSame(original, copy);
            Assert.Equal(text, copy.Text);
            Assert.Equal(startBase, copy.StartBase);
            Assert.Equal(endBase, copy.EndBase);
        }

        [Fact(DisplayName = "Test of the Copy method creates deep copy")]
        public void CopyCreatesDeepCopyTest()
        {
            PhoneticRun original = new PhoneticRun("Original", 0, 5);
            PhoneticRun copy = original.Copy();
            copy.Text = "Modified";
            copy.StartBase = 10;
            copy.EndBase = 20;
            Assert.Equal("Original", original.Text);
            Assert.Equal(0u, original.StartBase);
            Assert.Equal(5u, original.EndBase);
            Assert.Equal("Modified", copy.Text);
            Assert.Equal(10u, copy.StartBase);
            Assert.Equal(20u, copy.EndBase);
        }

        [Fact(DisplayName = "Test of the Copy method independence")]
        public void CopyIndependenceTest()
        {
            PhoneticRun original = new PhoneticRun("テスト", 0, 5);
            PhoneticRun copy = original.Copy();
            copy.StartBase = 100;
            Assert.Equal(0u, original.StartBase);
            Assert.Equal(100u, copy.StartBase);
        }

        [Fact(DisplayName = "Test of the PhoneticType enum values")]
        public void PhoneticTypeEnumValuesTest()
        {
            Assert.Equal(0, (int)PhoneticRun.PhoneticType.HalfwidthKatakana);
            Assert.Equal(1, (int)PhoneticRun.PhoneticType.FullwidthKatakana);
            Assert.Equal(2, (int)PhoneticRun.PhoneticType.Hiragana);
            Assert.Equal(3, (int)PhoneticRun.PhoneticType.NoConversion);
        }

        [Fact(DisplayName = "Test of the PhoneticAlignment enum values")]
        public void PhoneticAlignmentEnumValuesTest()
        {
            Assert.Equal(0, (int)PhoneticRun.PhoneticAlignment.NoControl);
            Assert.Equal(1, (int)PhoneticRun.PhoneticAlignment.Left);
            Assert.Equal(2, (int)PhoneticRun.PhoneticAlignment.Center);
            Assert.Equal(3, (int)PhoneticRun.PhoneticAlignment.Distributed);
        }
    }
}
