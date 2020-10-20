using OfficeOpenXml.Drawing.Vml;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Controls
{
    public abstract class ExcelControlBase : ExcelDrawing
    {
        ExcelVmlDrawingControl _vml;
        XmlHelper _ctrlProp;
        internal ExcelControlBase(ExcelDrawings drawings, XmlNode node, XmlNode vmlTopNode, XmlDocument ctrlPropXml, ExcelGroupShape parent = null) : base(drawings, node, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr", parent)
        {
            _vml = new ExcelVmlDrawingControl(vmlTopNode, NameSpaceManager);
            ControlPropertiesXml = ctrlPropXml;
            _ctrlProp = XmlHelperFactory.Create(NameSpaceManager, ctrlPropXml);
        }
        public XmlDocument ControlPropertiesXml { get; private set; }
        public abstract eControlType ControlType
        {
            get;
        }
    }

    public enum eControlType
    {
        Button,
        Combobox
    }
}
