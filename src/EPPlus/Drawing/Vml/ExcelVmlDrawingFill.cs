using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml
{
    public class ExcelVmlDrawingFill : XmlHelper
    {
        internal ExcelVmlDrawingFill(XmlNode topNode, XmlNamespaceManager ns) :
            base(ns, topNode)
        {
            SchemaNodeOrder = new string[] { "fill", "stroke", "shadow", "path", "textbox", "ClientData", "MoveWithCells", "SizeWithCells", "Anchor", "Locked", "AutoFill", "LockText", "TextHAlign", "TextVAlign", "Row", "Column", "Visible" };
        }
        public eVmlFillType Style 
        {
            get
            {
                return GetXmlNodeString("type").ToEnum<eVmlFillType>();
            }
            set
            {
                SetXmlNodeString("type", value.ToEnum<eVmlFillType>());
            }
        }
    }
}
