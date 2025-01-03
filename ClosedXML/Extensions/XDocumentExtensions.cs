
using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace ClosedXML.Excel
{
    internal static class XDocumentExtensions
    {
        public static XDocument Load(Stream stream)
        {
            using var reader = XmlReader.Create(stream);
            try
            {
                return XDocument.Load(reader);
            }
            catch (XmlException)
            {
                return null;
            }
        }
    }
}
