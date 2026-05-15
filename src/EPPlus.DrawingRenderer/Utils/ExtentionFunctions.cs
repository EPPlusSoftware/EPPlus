using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace EPPlus.DrawingRenderer.Utils
{
    internal static class ExtentionFunctions
    {
        internal static bool IsElementWithName(this XmlReader xr, string name)
        {
            return xr.NodeType == XmlNodeType.Element && xr.LocalName == name;
        }
        internal static bool IsEndElementWithName(this XmlReader xr, string name)
        {
            return xr.NodeType == XmlNodeType.EndElement && xr.LocalName == name;
        }
    }
}
