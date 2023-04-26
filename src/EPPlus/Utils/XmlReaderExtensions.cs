using System.Xml;

namespace OfficeOpenXml.Utils
{
    internal static class XmlReaderExtensions
    {
        internal static bool IsElementWithName(this XmlReader xr, string name)
        {
            return xr.NodeType == XmlNodeType.Element && xr.Name == name;
        }
        internal static bool IsEndElementWithName(this XmlReader xr, string name)
        {
            return xr.NodeType == XmlNodeType.EndElement && xr.Name == name;
        }
    }
}
