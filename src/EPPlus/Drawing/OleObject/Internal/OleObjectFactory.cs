using System;
using System.IO;
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

        internal static ExcelOleObject CreateOleObject(ExcelDrawings drawings, XmlElement drawNode, string name, string olePath, ExcelOleObjectParameters parameters, string iconFilePath = null)
        {
            return new ExcelOleObject(drawings, drawNode, name, olePath, parameters, iconFilePath);
        }
        internal static ExcelOleObject CreateOleObject(ExcelDrawings drawings, XmlElement drawNode, string name, FileInfo oleInfo, ExcelOleObjectParameters parameters, FileInfo iconInfo = null)
        {
            return new ExcelOleObject(drawings, drawNode, name, oleInfo, parameters, iconInfo);
        }
        internal static ExcelOleObject CreateOleObject(ExcelDrawings drawings, XmlElement drawNode, string name, Stream oleStream, ExcelOleObjectParameters parameters, Stream iconStream = null)
        {
            return new ExcelOleObject(drawings, drawNode, name, oleStream, parameters, iconStream);
        }
    }
}
