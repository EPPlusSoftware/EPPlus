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
        public string Text 
        { 
            get
            {
                return GetXmlNodeString("v:textbox/d:div/d:font");
            }
            set
            {
                SetXmlNodeString("v:textbox/div/font", value);
            }
        }
    }
}
