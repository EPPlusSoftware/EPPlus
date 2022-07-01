using System.Xml;

namespace OfficeOpenXml.Utils.Extensions
{
    internal static class XmlExtensions
    {
        internal static XmlNode GetChildAtPosition(this XmlNode node, int index, XmlNodeType type = XmlNodeType.Element)
        {
            var i = 0;
            foreach (XmlNode c in node.ChildNodes)
            {
                if (c.NodeType == type)
                {
                    if (i == index)
                    {
                        return c;
                    }
                    i++;
                }
            }
            return null;
        }
    }
}
