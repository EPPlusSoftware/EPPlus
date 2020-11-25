/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    11/24/2020         EPPlus Software AB           Controls 
 *************************************************************************************************/
using OfficeOpenXml.Constants;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Xml;
namespace OfficeOpenXml.Drawing.Controls
{
    public abstract class ExcelControl : ExcelDrawing
    {
        protected ExcelVmlDrawingControl _vml;
        protected XmlHelper _ctrlProp;
        internal protected XmlHelper _vmlProp;
        internal ControlInternal _control;

        internal ExcelControl(ExcelDrawings drawings, XmlNode drawingNode, ControlInternal control, ZipPackagePart ctrlPropPart, XmlDocument ctrlPropXml, ExcelGroupShape parent = null) :
            base(drawings, drawingNode, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr", parent)
        {
            _control = control;
            _vml = (ExcelVmlDrawingControl)drawings.Worksheet.VmlDrawings[LegacySpId];
            _vmlProp = XmlHelperFactory.Create(_vml.NameSpaceManager, _vml.GetNode("x:ClientData"));
            ControlPropertiesXml = ctrlPropXml;
            ControlPropertiesPart = ctrlPropPart;
            ControlPropertiesUri = ctrlPropPart.Uri;
            _ctrlProp = XmlHelperFactory.Create(NameSpaceManager, ctrlPropXml.DocumentElement);
        }

        protected ExcelControl(ExcelDrawings drawings, XmlNode drawingNode) : base(drawings, drawingNode, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr")
        {
            var ws = drawings.Worksheet;

            //Drawing Xml
            XmlElement spElement = CreateShapeNode();
            spElement.InnerXml = ControlStartDrawingXml();
            CreateClientData();

            ControlPropertiesXml = new XmlDocument();
            ControlPropertiesXml.LoadXml(ControlStartControlPrXml());            
            int id= ws.SheetId;
            ControlPropertiesUri = GetNewUri(ws._package.ZipPackage, "/xl/ctrlProps/ctrlProp{0}.xml",ref id);
            ControlPropertiesPart = ws._package.ZipPackage.CreatePart(ControlPropertiesUri, ContentTypes.contentTypeControlProperties);
            var rel=ws.Part.CreateRelationship(ControlPropertiesUri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/ctrlProp");

            //Vml
            _vml=drawings.Worksheet.VmlDrawings.AddControl(this);
            _vmlProp = XmlHelperFactory.Create(_vml.NameSpaceManager, _vml.GetNode("x:ClientData"));

            //Control in worksheet xml
            XmlNode ctrlNode = ws.CreateControlNode();
            ((XmlElement)ws.TopNode).SetAttribute("xmlns:xdr", ExcelPackage.schemaSheetDrawings);   //Make sure the namespace exists
            ((XmlElement)ws.TopNode).SetAttribute("xmlns:x14", ExcelPackage.schemaMainX14);   //Make sure the namespace exists
            ((XmlElement)ws.TopNode).SetAttribute("xmlns:mc", ExcelPackage.schemaMarkupCompatibility);   //Make sure the namespace exists
            ctrlNode.InnerXml = GetControlStartWorksheetXml(rel.Id);
            _control = new ControlInternal(NameSpaceManager, ctrlNode.FirstChild.FirstChild.FirstChild);
            _ctrlProp = XmlHelperFactory.Create(NameSpaceManager, ControlPropertiesXml.DocumentElement);
        }
        private string GetControlStartWorksheetXml(string relId)
        {
            var sb = new StringBuilder();

            sb.Append($"<mc:AlternateContent><mc:Choice Requires=\"x14\"><control shapeId=\"{Id}\" r:id=\"{relId}\" name=\"\">");
            sb.Append("<controlPr defaultSize=\"0\" print=\"0\" autoFill=\"0\" autoPict=\"0\">");
            sb.Append("<anchor moveWithCells=\"1\" sizeWithCells=\"1\">");
            sb.Append($"<from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></from>");
            sb.Append($"<to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></to>");
            sb.Append("</anchor></controlPr></control></mc:Choice></mc:AlternateContent>");
            return sb.ToString();
        }
        private string ControlStartControlPrXml()
        {
            var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><formControlPr xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" {0} />";
            switch (ControlType)
            {
                case eControlType.Button:
                    return string.Format(xml, "objectType=\"Button\" lockText=\"1\"");
                case eControlType.CheckBox:
                    return string.Format(xml, "objectType=\"CheckBox\" lockText=\"1\" noThreeD=\"1\"");
                case eControlType.RadioButton:
                    return string.Format(xml, "objectType=\"Radio\" firstButton=\"1\" lockText=\"1\" noThreeD=\"1\"");
                case eControlType.DropDown:
                    return string.Format(xml, "objectType=\"Drop\" dropStyle=\"combo\" dx=\"22\" noThreeD=\"1\" sel=\"0\" val=\"0\"");
                case eControlType.ListBox:
                    return string.Format(xml, "objectType=\"List\" dx=\"22\" noThreeD=\"1\" sel=\"0\" val=\"0\"");
                case eControlType.Label:
                    return string.Format(xml, "objectType=\"Label\" lockText=\"1\"");
                case eControlType.ScrollBar:
                    return string.Format(xml, "objectType=\"Scroll\" dx=\"22\" max=\"100\" page=\"10\" val=\"0\"");
                case eControlType.SpinButton:
                    return string.Format(xml, "objectType=\"Spin\" dx=\"22\" fmlaLink=\"$A$1\" max=\"30000\" page=\"10\" val=\"0\"");
                case eControlType.GroupBox:
                    return string.Format(xml, "objectType=\"GBox\" noThreeD=\"1\"");
                default:
                    throw new NotImplementedException();
            }
        }

        private string ControlStartDrawingXml()
        {
            StringBuilder xml = new StringBuilder();
            xml.Append($"<xdr:nvSpPr><xdr:cNvPr hidden=\"1\" name=\"\" id=\"{_id}\"><a:extLst><a:ext uri=\"{{63B3BB69-23CF-44E3-9099-C40C66FF867C}}\"><a14:compatExt spid=\"_x0000_s{_id}\"/></a:ext><a:ext uri=\"{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}\"><a16:creationId id=\"{{00000000-0008-0000-0000-000001040000}}\" xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr>");
            xml.Append($"<xdr:spPr bwMode=\"auto\"><a:xfrm><a:off y=\"0\" x=\"0\"/><a:ext cy=\"0\" cx=\"0\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>");
            switch (ControlType)
            {
                case eControlType.Button:
                    xml.Append($"<a:noFill/><a:ln w=\"9525\"><a:miter lim=\"800000\"/><a:headEnd/><a:tailEnd/></a:ln>");
                    break;
                case eControlType.CheckBox:
                case eControlType.RadioButton:
                    xml.Append($"<a:noFill/><a:ln><a:noFill/></a:ln><a:extLst><a:ext uri=\"{{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}}\"><a14:hiddenFill><a:solidFill><a:srgbClr val=\"FFFFFF\" mc:Ignorable=\"a14\" a14:legacySpreadsheetColorIndex=\"65\"/></a:solidFill></a14:hiddenFill></a:ext><a:ext uri=\"{{91240B29-F687-4F45-9708-019B960494DF}}\"><a14:hiddenLine w=\"9525\"><a:solidFill><a:srgbClr val=\"000000\" mc:Ignorable=\"a14\" a14:legacySpreadsheetColorIndex=\"64\"/></a:solidFill><a:miter lim=\"800000\"/><a:headEnd/><a:tailEnd/></a14:hiddenLine></a:ext></a:extLst>");
                    break;
                case eControlType.ListBox:
                    xml.Append("<a:noFill/><a:ln><a:noFill/></a:ln><a:extLst><a:ext uri=\"{{91240B29-F687-4F45-9708-019B960494DF}}\"><a14:hiddenLine w=\"9525\"><a:noFill/><a:miter lim=\"800000\"/><a:headEnd/><a:tailEnd/></a14:hiddenLine></a:ext></a:extLst>");
                    break;
            }
            xml.Append("</xdr:spPr>");
            if(this is ExcelControlWithText)
            {
                xml.Append($"<xdr:txBody><a:bodyPr upright=\"1\" anchor=\"{GetDrawingAnchor()}\" bIns=\"27432\" rIns=\"27432\" tIns=\"27432\" lIns=\"27432\" wrap=\"square\" vertOverflow=\"clip\"/>" +
                    $"<a:lstStyle/>" +
                    $"<a:p>{GetrPr(ControlType)}" +
                    $"<a:r><a:rPr lang=\"en-US\" sz=\"{GetFontSize()}\" baseline=\"0\" strike=\"noStrike\" u=\"none\" i=\"0\" b=\"0\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill><a:latin typeface=\"{GetFontName()}\"/><a:cs typeface=\"{GetFontName()}\"/></a:rPr><a:t></a:t></a:r></a:p></xdr:txBody>");
            }
            return xml.ToString();
        }

        private string GetFontName()
        {
            if (ControlType == eControlType.Button)
            {
                return "Calibri";
            }
            else
            {
                return "Segoe UI";
            }
        }

        private string GetFontSize()
        {
            if(ControlType==eControlType.Button)
            {
                return "1100";
            }
            else
            {
                return "800";
            }
        }

        private string GetDrawingAnchor()
        {
            if(ControlType==eControlType.GroupBox)
            {
                return "t";
            }
            else
            {
                return "ctr";
            }
        }

        private static string GetrPr(eControlType controlType)
        {
            switch (controlType)
            {
                case eControlType.Button:
                    return "<a:pPr rtl=\"0\" algn=\"ctr\"><a:defRPr sz=\"1000\"/></a:pPr>";
                case eControlType.CheckBox:
                case eControlType.RadioButton:
                case eControlType.Label:
                    return "<a:pPr rtl=\"0\" algn=\"l\"><a:defRPr sz=\"1000\"/></a:pPr>"; 
                default:
                    return "<a:pPr rtl=\"0\" algn=\"l\"><a:defRPr sz=\"1000\"/></a:pPr>";                    
            }
        }

        private XmlNode GetVmlNode(ExcelVmlDrawingCollection vmlDrawings)
        {
            return vmlDrawings.FirstOrDefault(x => x.Id == LegacySpId)?.TopNode;
        }

        public XmlDocument ControlPropertiesXml { get; private set; }
        internal ZipPackagePart ControlPropertiesPart { get; private set; }
        internal Uri ControlPropertiesUri { get; private set; }
        public abstract eControlType ControlType
        {
            get;
        }
        internal string ControlTypeString
        {
            get
            {
                switch(ControlType)
                {
                    case eControlType.GroupBox:
                        return "GBox";
                    case eControlType.CheckBox:
                        return "Checkbox";
                    case eControlType.RadioButton:
                        return "Radio";
                    case eControlType.DropDown:
                        return "Drop";
                    case eControlType.ListBox:
                        return "List";
                    case eControlType.SpinButton:
                        return "Spin";
                    case eControlType.ScrollBar:
                        return "Scroll";
                    default:
                        return ControlType.ToString();
                }
            }
        }
        internal string LegacySpId
        {
            get
            {
                return GetXmlNodeString($"xdr:sp/xdr:nvSpPr/xdr:cNvPr/a:extLst/a:ext[@uri='{ExtLstUris.LegacyObjectWrapperUri}']/a14:compatExt/@spid");
            }
            set
            {
                var node= GetNode("xdr:sp/xdr:nvSpPr/xdr:cNvPr");
                var extHelper = XmlHelperFactory.Create(NameSpaceManager, node);
                var extNode= extHelper.GetOrCreateExtLstSubNode(ExtLstUris.LegacyObjectWrapperUri, "a14");
                if (extNode.InnerXml == "")
                {
                    extNode.InnerXml = $"<a14:compatExt/>";
                }
                ((XmlElement)extNode.FirstChild).SetAttribute("spid", value);

            }
        }
        /// <summary>
        /// The name of the control
        /// </summary>
        public override string Name
        {
            get
            {
                return _control.Name;
            }
            set
            {
                _control.Name=value;
                base.Name = value;
            }
        }
        /// <summary>
        /// Gets or sets the alternative text for the control.
        /// </summary>
        public string AlternativeText
        {
            get
            {
                return _control.AlternativeText;
            }
            set
            {
                _control.AlternativeText = value;
            }
        }
        /// <summary>
        /// Gets or sets the macro function assigned.
        /// </summary>
        public string Macro
        {
            get
            {
                return _control.Macro;
            }
            set
            {
                _control.Macro = value;
                _vmlProp.SetXmlNodeString("x:FmlaMacro", value);
            }
        }

        internal string GetVmlAnchorValue()
        {
            return string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}",
                From.Column, From.ColumnOff / EMU_PER_PIXEL,
                From.Row, From.RowOff / EMU_PER_PIXEL,
                To.Column, To.ColumnOff / EMU_PER_PIXEL,
                To.Row, To.RowOff / EMU_PER_PIXEL);
        }

        /// <summary>
        /// The object is printed when the document is printed.
        /// </summary>
        public override bool Print
        {
            get
            {
                return _control.Print;
            }
            set
            {
                _control.Print = value;
                base.Print = value;
            }
        }

        /// <summary>
        /// The object is locked when the sheet is protected..
        /// </summary>
        public override bool Locked
        {
            get
            {
                return _control.Locked;
            }
            set
            {
                _control.Locked = value;
                base.Locked = value;
            }
        }
        /// <summary>
        /// If the controls fill formatting is provided automatically
        /// </summary>
        public bool AutoFill
        {
            get { return _control.AutoFill; }
            set { _control.AutoFill = value; }
        }

        /// <summary>
        /// If the controls size is formatted automatically.
        /// </summary>
        public bool AutoPict
        {
            get { return _control.AutoPict; }
            set { _control.AutoPict = value; }
        }

        /// <summary>
        /// Returns true if the object is at its default size.
        /// </summary>
        public bool DefaultSize
        {
            get { return _control.DefaultSize; }
            set { _control.DefaultSize = value; }
        }

        /// <summary>
        /// If true, the object is allowed to run an attached macro
        /// </summary>
        public bool Disabled
        {
            get { return _control.Disabled; }
            set { _control.Disabled = value; }
        }
        /// <summary>
        /// If the control has 3D effects enabled.
        /// </summary>
        public bool ThreeDEffects
        {
            get
            {
                var b = _ctrlProp.GetXmlNodeBoolNullable("@noThreeD2");
                if (b.HasValue == false)
                {
                    return _ctrlProp.GetXmlNodeBool("@noThreeD") == false;
                }
                else
                {
                    return _ctrlProp.GetXmlNodeBool("@noThreeD2") == false;
                }
            }
            set
            {
                var b = _ctrlProp.GetXmlNodeBoolNullable("@noThreeD2");
                if (b.HasValue)
                {
                    _ctrlProp.SetXmlNodeBool("@noThreeD2", value == false);   //can be used for lists and drop-downs.
                }
                else
                {
                    _ctrlProp.SetXmlNodeBool("@noThreeD", value == false);
                }

                var xmlAttr = (ControlType == eControlType.DropDown || ControlType == eControlType.ListBox) ? "x:NoThreeD2" : "x:NoThreeD";
                if (value)
                {
                    _vmlProp.CreateNode(xmlAttr);
                }
                else
                {
                    _vmlProp.DeleteNode(xmlAttr);
                }
            }
        }
        /// <summary>
        /// Gets or sets the address to the cell that is linked to the control. 
        /// </summary>
        public ExcelAddressBase LinkedCell
        {
            get
            {
                if (ControlType == eControlType.Label ||
                   ControlType == eControlType.Button ||
                   ControlType == eControlType.GroupBox)
                {
                    return FmlaTxbx;
                }
                else
                {
                    return FmlaLink;
                }

            }
            set
            {
                if (ControlType == eControlType.Label ||
                   ControlType == eControlType.Button ||
                   ControlType == eControlType.GroupBox)
                {
                    FmlaTxbx = value;
                }
                else
                {
                    FmlaLink = value;
                }
            }
        }
        #region Shared Properties
        internal protected ExcelAddressBase FmlaLink
        {
            get
            {
                var range = _ctrlProp.GetXmlNodeString("@fmlaLink");
                if (ExcelAddressBase.IsValidAddress(range))
                {
                    return new ExcelAddressBase(range);
                }
                return null;
            }
            set
            {
                if (value == null)
                {
                    _ctrlProp.DeleteNode("@fmlaLink");
                    _vmlProp.DeleteNode("x:FmlaLink");
                }
                if (value.WorkSheetName.Equals(_drawings.Worksheet.Name, StringComparison.CurrentCultureIgnoreCase))
                {
                    _ctrlProp.SetXmlNodeString("@fmlaLink", value.Address);
                    _vmlProp.SetXmlNodeString("x:FmlaLink", value.Address);
                }
                else
                {
                    _ctrlProp.SetXmlNodeString("@fmlaLink", value.FullAddress);
                    _vmlProp.SetXmlNodeString("x:FmlaLink", value.FullAddress);
                }
            }
        }
        /// <summary>
        /// The source data cell that the control object's data is linked to.
        /// </summary>
        internal ExcelAddressBase FmlaTxbx
        {
            get
            {
                var range = _ctrlProp.GetXmlNodeString("@fmlaTxbx");
                if (ExcelAddressBase.IsValidAddress(range))
                {
                    return new ExcelAddressBase(range);
                }
                return null;
            }
            set
            {
                if (value == null)
                {
                    _ctrlProp.DeleteNode("@fmlaTxbx");
                    _vmlProp.DeleteNode("x:FmlaTxbx");
                }
                if (value.WorkSheetName.Equals(_drawings.Worksheet.Name, StringComparison.CurrentCultureIgnoreCase))
                {
                    _ctrlProp.SetXmlNodeString("@fmlaTxbx", value.Address);
                    _vmlProp.SetXmlNodeString("x:FmlaTxbx", value.Address);
                }
                else
                {
                    _ctrlProp.SetXmlNodeString("@fmlaTxbx", value.FullAddress);
                    _vmlProp.SetXmlNodeString("x:FmlaTxbx", value.FullAddress);
                }
            }
        }
        internal protected ExcelAddressBase LinkedGroup
        {
            get
            {
                var range = _ctrlProp.GetXmlNodeString("@fmlaGroup");
                if (ExcelAddressBase.IsValidAddress(range))
                {
                    return new ExcelAddressBase(range);
                }
                return null;
            }
            set
            {
                if (value == null)
                {
                    _ctrlProp.DeleteNode("@fmlaGroup");
                    _vmlProp.DeleteNode("x:FmlaGroup");
                }
                if (value.WorkSheetName.Equals(_drawings.Worksheet.Name, StringComparison.CurrentCultureIgnoreCase))
                {
                    _ctrlProp.SetXmlNodeString("@fmlaGroup", value.Address);
                    _vmlProp.SetXmlNodeString("x:FmlaGroup", value.Address);
                }
                else
                {
                    _ctrlProp.SetXmlNodeString("@fmlaGroup", value.FullAddress);
                    _vmlProp.SetXmlNodeString("x:FmlaGroup", value.FullAddress);
                }
            }
        }
        public override eDrawingType DrawingType
        {
            get
            {
                return eDrawingType.Control;
            }
        }

        internal virtual void UpdateXml()
        {
            _control.From.Row = From.Row;
            _control.From.RowOff = From.RowOff;
            _control.From.Column = From.Column;
            _control.From.ColumnOff = From.ColumnOff;
            
            _control.To.Row = To.Row;
            _control.To.RowOff = To.RowOff;
            _control.To.Column = To.Column;
            _control.To.ColumnOff = To.ColumnOff;

            _control.MoveWithCells = EditAs != eEditAs.Absolute;
            _control.SizeWithCells = EditAs == eEditAs.TwoCell;

            _vml.Anchor = GetVmlAnchorValue();
        }

        internal static eEditAs GetControlEditAs(eControlType controlType)
        {
            switch(controlType)
            {
                case eControlType.Button:
                    return eEditAs.Absolute;
                case eControlType.Label:
                    return eEditAs.TwoCell;
                default:
                    return eEditAs.OneCell;
            }
        }

        #endregion
    }
}
