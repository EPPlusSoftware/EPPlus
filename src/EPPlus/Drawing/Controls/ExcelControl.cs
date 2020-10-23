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
using System;
using System.Linq;
using System.Xml;
namespace OfficeOpenXml.Drawing.Controls
{
    public abstract class ExcelControl : ExcelDrawing
    {
        protected ExcelVmlDrawingControl _vml;
        protected XmlHelper _ctrlProp;
        internal ControlInternal _control;
        private ZipPackageRelationship _rel;

        internal ExcelControl(ExcelDrawings drawings, XmlNode drawingNode, ControlInternal control, ZipPackageRelationship rel, XmlDocument ctrlPropXml, ExcelGroupShape parent = null) : 
            base(drawings, drawingNode, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr", parent)
        {
            _control = control;
            _rel = rel;
            //VmlNode vmlTopNode=GetVmlNode(drawings.Worksheet.VmlDrawingsComments);
            var _vml = drawings.Worksheet.VmlDrawings[LegacySpId];
            ControlPropertiesXml = ctrlPropXml;
            //_vml = new ExcelVmlDrawingControl(vmlTopNode, NameSpaceManager);
            _ctrlProp = XmlHelperFactory.Create(NameSpaceManager, ctrlPropXml.DocumentElement);
        }

        private XmlNode GetVmlNode(ExcelVmlDrawingCollection vmlDrawings)
        {
            return vmlDrawings.FirstOrDefault(x=>x.Id==LegacySpId)?.TopNode;            
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
        public string Name
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

        public bool AutoFill
        {
            get { return _control.AutoFill; }
            set { _control.AutoFill = value; }
        }

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
    }

}
