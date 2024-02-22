using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml
{
    /// <summary>
    /// Base class for vml form controls
    /// </summary>
    public class ExcelVmlDrawingControl : ExcelVmlDrawingBase
    {
        ExcelWorksheet _ws;
        internal ExcelVmlDrawingControl(ExcelWorksheet ws, XmlNode topNode, XmlNamespaceManager ns) : base(topNode, ns)
        {
            _ws = ws;
        }
        /// <summary>
        /// The String
        /// </summary>
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
        internal ExcelVmlDrawingFill _fill = null;
        internal ExcelVmlDrawingFill GetFill()
        {
            if (_fill == null)
            {
                _fill = new ExcelVmlDrawingFill(_ws.Drawings, NameSpaceManager, TopNode, SchemaNodeOrder);
            }
            return _fill;
        }
    }
}
