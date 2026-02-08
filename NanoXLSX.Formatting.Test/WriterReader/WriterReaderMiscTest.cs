using NanoXLSX.Extensions;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Reflection;
using Xunit;

namespace NanoXLSX.Formatting.Test.WriterReader
{
    /// <summary>
    /// Misc tests for the reader or writer functionalities
    /// </summary>
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
    public class WriterReaderMiscTest : IDisposable
    {
        public WriterReaderMiscTest()
        {
            TestUtils.InitializePlugIns();
        }

        [Fact(DisplayName = "Test of the expected error handling of the reader plugin on invalid data (coverage)")]
        public void FailingReadInvalidDataTest()
        {
            // Note: The referenced (embedded) file contains an invalid XML document
            Stream stream = TestUtils.GetResource("invalid_sharedStrings2.xlsx");
            Assert.Throws<Exceptions.IOException>(() => WorkbookReader.Load(stream));
        }

        [Fact(DisplayName = "Test of the reader plugin on orphan phonetic properties")]
        public void ReadOrphanPhoneticPropertiesTest()
        {
            // Note: The referenced (embedded) file contains an XML document with rare conditions of phonetic properties without run
            Stream stream = TestUtils.GetResource("phonetic_properties.xlsx");
            Workbook workbook = WorkbookReader.Load(stream);
            Assert.NotNull(workbook);
            Cell cell = workbook.Worksheets[0].Cells["A1"];
            Assert.IsType<FormattedText>(cell.Value);
            Assert.NotNull(((FormattedText)cell.Value).PhoneticProperties);
            Assert.NotEmpty(cell.Value.ToString());
        }

        [Fact(DisplayName = "Test of the reader plugin on orphan phonetic runs")]
        public void ReadOrphanPhoneticRunTest()
        {
            // Note: The referenced (embedded) file contains an XML document with rare conditions of phonetic runs without text run
            Stream stream = TestUtils.GetResource("phonetic_run.xlsx");
            Workbook workbook = WorkbookReader.Load(stream);
            Assert.NotNull(workbook);
            Cell cell = workbook.Worksheets[0].Cells["A1"];
            Assert.IsType<FormattedText>(cell.Value);
            Assert.NotNull(((FormattedText)cell.Value).PhoneticRuns);
            Assert.NotEmpty(cell.Value.ToString());
        }

        [Fact(DisplayName = "Test of the reader plugin on orphan phonetic runs and options")]
        public void ReadOrphanPhoneticRunTest2()
        {
            // Note: The referenced (embedded) file contains an XML document with rare conditions of phonetic runs without text run
            Stream stream = TestUtils.GetResource("phonetic_run.xlsx");
            ReaderOptions options = new ReaderOptions();
            options.EnforcePhoneticCharacterImport = true;
            Workbook workbook = WorkbookReader.Load(stream, options);
            Assert.NotNull(workbook);
            Cell cell = workbook.Worksheets[0].Cells["A1"];
            Assert.IsType<FormattedText>(cell.Value);
            Assert.NotNull(((FormattedText)cell.Value).PhoneticRuns);
            Assert.NotEmpty(cell.Value.ToString());
        }

        [Fact(DisplayName = "Test of the reader plugin on handling default font properties (normally not written)")]
        public void ReadDefaultFontPropertiesTest()
        {
            // Note: The referenced (embedded) file contains an XML document with rare conditions of font properties:
            //       - scheme = none; Is usually minor by default
            //       - underline value = single; value is normally omitted and implicitly interpreted if not present
            Stream stream = TestUtils.GetResource("font_run.xlsx");
            Workbook workbook = WorkbookReader.Load(stream);
            Assert.NotNull(workbook);
            Cell cell = workbook.Worksheets[0].Cells["A1"];
            Assert.IsType<FormattedText>(cell.Value);
            TextRun textRun = ((FormattedText)cell.Value).Runs[0];
            Assert.Equal(Font.SchemeValue.None, textRun.FontStyle.Scheme);
            Assert.Equal(Font.UnderlineValue.Single, textRun.FontStyle.Underline);
        }

        [Fact(DisplayName = "Test of the in-line reader capabilities of the reader plugin")]
        public void InlineReaderPluginTest()
        {
            Type readerType = typeof(InlineSharedStringReader);
            string setupMethodName = nameof(SharedStringTestData);
            string expectedReferenceValue = "TestValue";
            string pluginUuid = PlugInUUID.SharedStringsInlineReader;
            Workbook wb = new Workbook("sheet1");
            var setupMethod = typeof(WriterReaderMiscTest).GetMethod(setupMethodName,
            BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            setupMethod.Invoke(null, new object[] { wb, expectedReferenceValue });

            List<Type> plugins = new List<Type>
            {
                readerType
            };
            PlugInLoader.InjectPlugins(plugins);

            wb.CurrentWorksheet.AddCell(expectedReferenceValue, "A1"); // Write reference value
            using (MemoryStream ms = new MemoryStream())
            {
                wb.SaveAsStream(ms, true);
                ms.Position = 0;
                Workbook wb2 = WorkbookReader.Load(ms);
                Assert.NotNull(wb2);
                string singleValue = wb2.AuxiliaryData.GetData<string>(pluginUuid, 0);
                Assert.Equal(expectedReferenceValue, wb2.AuxiliaryData.GetData<string>(pluginUuid, 0));
            }

        }

        public void Dispose()
        {
            TestUtils.DisposePlugIns();
        }

        public static void SharedStringTestData(Workbook wb, string expectedValue)
        {
            wb.CurrentWorksheet.AddNextCell(expectedValue);
        }

        [NanoXlsxQueuePlugIn(PlugInUUID = "SharedStringReaderPlugIn1", QueueUUID = PlugInUUID.SharedStringsInlineReader)]
        public class InlineSharedStringReader : IPluginInlineReader
        {
            private const string TEST_NODE = "t";
            private MemoryStream stream;
            public Workbook Workbook { get; set; }
            public IOptions Options { get; set; }
            [ExcludeFromCodeCoverage]
            public Action<MemoryStream, Workbook, string, IOptions, int?> InlinePluginHandler { get; set; }

            public void Execute()
            {
                string testValue = TestUtils.ReadFirstNodeValue(stream, TEST_NODE);
                Workbook.AuxiliaryData.SetData(PlugInUUID.SharedStringsInlineReader, 0, testValue, true);
                this.stream.Position = 0;
            }

            public void Init(MemoryStream stream, Workbook workbook, IOptions readerOptions, int? index = null)
            {
                this.stream = stream;
                this.stream.Position = 0;
                this.Workbook = workbook;
                this.Options = readerOptions;
            }
        }
    }
}
