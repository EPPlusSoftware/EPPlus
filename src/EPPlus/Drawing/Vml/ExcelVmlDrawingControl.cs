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
        /// <summary>
        /// Item height for an individual item
        /// </summary>
        internal int? Dx
        {
            get
            {
                return GetXmlNodeIntNull("x:ClientData/x:Dx");
            }
            set
            {
                SetXmlNodeInt("x:ClientData/x:Dx", value);
            }
        }
        /// <summary>
        /// Number of items in a listbox.
        /// </summary>
        internal int? Page
        {
            get
            {
                return GetXmlNodeIntNull("x:ClientData/x:Page");
            }
            set
            {
                SetXmlNodeInt("x:ClientData/x:Page", value);
            }
        }

    }
}
