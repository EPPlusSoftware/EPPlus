using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.OleObject
{
    internal static class OleObjectFactory
    {
        internal static ExcelDrawing GetOleObject(ExcelDrawings drawings, XmlElement drawNode, OleObjectInternal oleObject, ExcelGroupShape parent)
        {
            XmlNode node;
            if (parent == null)
            {
                node = drawNode.ParentNode;
            }
            else
            {
                node = drawNode;
            }
            return new ExcelOleObject(drawings, node, oleObject, parent);
        }

        internal static ExcelOleObject CreateOleObject(ExcelDrawings drawings, XmlElement drawNode, string filepath, bool linkToFile, bool displayAsIcon, string iconFilePath)
        {
            return new ExcelOleObject(drawings, drawNode, filepath, linkToFile, displayAsIcon, iconFilePath);
        }
    }
}
