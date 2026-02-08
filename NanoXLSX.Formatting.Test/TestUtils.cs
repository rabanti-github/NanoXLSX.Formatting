using NanoXLSX.Extensions;
using NanoXLSX.Registry;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Reflection;
using System.Xml;

namespace NanoXLSX.Formatting.Test
{
    [ExcludeFromCodeCoverage]
    internal static class TestUtils
    {
        private const string ASSEMBLY_RESOURCE_NAMESPACE = "NanoXLSX.Formatting.Test"; // Change this on refactoring

        internal static Workbook SaveAndReadWorkbook(Workbook workbook)
        {
            MemoryStream stream = new MemoryStream();
            workbook.SaveAsStream(stream, true);
            stream.Position = 0;
            Workbook loadedWorkbook = WorkbookReader.Load(stream);
            return loadedWorkbook;
        }

        internal static bool IsEqual(double? a, double? b, double tolerance = 0.0001)
        {
            if (a == null && b == null)
            {
                return true;
            }
            else if (a == null || b == null)
            {
                return false;
            }
            return Math.Abs(a.Value - b.Value) <= tolerance;
        }

        public static Stream GetResource(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return null;
            }
            Assembly assembly = Assembly.GetExecutingAssembly();
            string resourceName = $"{ASSEMBLY_RESOURCE_NAMESPACE}.Resources.{path}";
            try
            {
                return assembly.GetManifestResourceStream(resourceName);
            }
            catch
            {
                return null;
            }
        }

        internal static void InitializePlugIns()
        {
            PlugInLoader.DisposePlugins();
            PlugInLoader.InjectPlugins(new System.Collections.Generic.List<System.Type>
            {
                 typeof(Internal.Readers.FormattedSharedStringsReader),
                 typeof(Internal.Readers.SharedStringsReplacer),
                 typeof(Internal.FontIdResolver)
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

        /// <summary>
        /// Reads the first inner text value of the given node name from an XML stream.
        /// </summary>
        /// <param name="stream">The XML content as a readable stream.</param>
        /// <param name="nodeName">The name of the XML node to read.</param>
        /// <returns>The first inner text value of the node, or null if not found.</returns>
        internal static string ReadFirstNodeValue(Stream stream, string nodeName)
        {
            if (stream == null)
            {
                throw new ArgumentNullException(nameof(stream));
            }
            if (string.IsNullOrEmpty(nodeName))
            {
                throw new ArgumentException("Node name must be specified.", nameof(nodeName));
            }
            using (var reader = XmlReader.Create(stream, new XmlReaderSettings
            {
                IgnoreComments = true,
                IgnoreWhitespace = true,
                DtdProcessing = DtdProcessing.Ignore
            }))
            {
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == nodeName)
                    {
                        // ReadElementContentAsString moves past the end element automatically
                        return reader.ReadElementContentAsString();
                    }
                }
            }

            return null; // not found
        }
    }
}
