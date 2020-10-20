using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml
{
    public class ExcelVmlDrawingControl : ExcelVmlDrawingBase
    {
        internal ExcelVmlDrawingControl(XmlNode topNode, XmlNamespaceManager ns) : base(topNode, ns)
        {
        }
    }
}
