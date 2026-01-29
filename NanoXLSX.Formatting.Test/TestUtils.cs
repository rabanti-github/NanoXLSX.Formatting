using NanoXLSX.Extensions;
using NanoXLSX.Registry;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NanoXLSX.Formatting.Test
{
    internal static class TestUtils
    {
        internal static Workbook SaveAndReadWorkbook(Workbook workbook)
        {
            MemoryStream stream = new MemoryStream();
            workbook.SaveAsStream(stream, true);
            stream.Position = 0;
            Workbook loadedWorkbook = WorkbookReader.Load(stream);
            return loadedWorkbook;
        }

        internal static void InitializePlugIns()
        {
            PlugInLoader.DisposePlugins();
            PlugInLoader.InjectPlugins(new System.Collections.Generic.List<System.Type>
            {
                 typeof(Internal.Readers.FormattedSharedStringsReader),
                 typeof(Internal.Readers.SharedStringsReplacer)
            });
        }

        internal static void DisposePlugIns()
        {
            PlugInLoader.DisposePlugins();
        }

        internal static FormattedText GetFormattedText(List<Tuple<string, Font>> runs)
        {
            FormattedText text = new FormattedText();
            foreach (Tuple<string, Font> run in runs)
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
    }
}
