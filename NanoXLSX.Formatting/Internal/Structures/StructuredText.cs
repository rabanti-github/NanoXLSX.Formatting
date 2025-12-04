using System;
using System.Collections.Generic;
using System.Text;
using NanoXLSX.Interfaces;
using NanoXLSX.Utils.Xml;

namespace NanoXLSX.Internal.Structures
{
    /// <summary>
    /// Class to handle formatted texts, when writing into the shared string table 
    /// </summary>
    /// \remark <remarks>This class is only for internal use. Use the high level API (e.g. class Worksheet) to manipulate data and create Excel files</remarks>
    public class StructuredText : IFormattableText
    {
        /// <summary>
        /// Get the XmlElement (interface implementation)
        /// </summary>
        /// <returns>XmlElement instance</returns>
        public XmlElement GetXmlElement()
        {
            throw new NotImplementedException();
        }
    }
}
