/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    10/21/2020         EPPlus Software AB           Controls 
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Linq;
using System.Xml;
namespace OfficeOpenXml.Drawing.Controls
{
    public abstract class ExcelControl : ExcelDrawing
    {
        protected ExcelVmlDrawingControl _vml;
        protected XmlHelper _ctrlProp;
        protected XmlHelper _vmlProp;
        internal ControlInternal _control;
        private ZipPackageRelationship _rel;

        internal ExcelControl(ExcelDrawings drawings, XmlNode drawingNode, ControlInternal control, ZipPackageRelationship rel, XmlDocument ctrlPropXml, ExcelGroupShape parent = null) :
            base(drawings, drawingNode, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr", parent)
        {
            _control = control;
            _rel = rel;
            var _vml = drawings.Worksheet.VmlDrawings[LegacySpId];
            _vmlProp = XmlHelperFactory.Create(NameSpaceManager, _vml.GetNode("x:ClientData"));
            ControlPropertiesXml = ctrlPropXml;
            _ctrlProp = XmlHelperFactory.Create(NameSpaceManager, ctrlPropXml.DocumentElement);
        }

        private XmlNode GetVmlNode(ExcelVmlDrawingCollection vmlDrawings)
        {
            return vmlDrawings.FirstOrDefault(x => x.Id == LegacySpId)?.TopNode;
        }

        public XmlDocument ControlPropertiesXml { get; private set; }
        public abstract eControlType ControlType
        {
            get;
        }
        internal string LegacySpId
        {
            get
            {
                return GetXmlNodeString("xdr:sp/xdr:nvSpPr/xdr:cNvPr/a:extLst/a:ext[@uri='{63B3BB69-23CF-44E3-9099-C40C66FF867C}']/a14:compatExt/@spid");
            }
            set
            {

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
        /// Horizontal text alignment. Not used in Excel 2010- , so internal for now
        /// </summary>
        internal eHorizontalAlignmentControl HorizontalTextAlignment
        {
            get
            {
                return _ctrlProp.GetXmlNodeString("textHAlign").ToEnum(eHorizontalAlignmentControl.Left);
            }
            set
            {
                _ctrlProp.SetXmlNodeString("textHAlign", value.ToEnumString());
                _vmlProp.SetXmlNodeString("x:TextHAlign",value.ToString());
            }
        }
        /// <summary>
        /// Vertical text alignment. Not used in Excel 2010-
        /// </summary>
        internal eVerticalAlignmentControl VerticalTextAlignment
        {
            get
            {
                return _ctrlProp.GetXmlNodeString("textVAlign").ToEnum(eVerticalAlignmentControl.Top);
            }
            set
            {
                _ctrlProp.SetXmlNodeString("textVAlign", value.ToEnumString());
                _vmlProp.SetXmlNodeString("x:TextVAlign", value.ToString());
            }
        }
        #region Shared Properties
        internal protected ExcelAddressBase LinkedCellBase
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
        #endregion
    }
}
