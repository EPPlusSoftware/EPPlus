/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
  02/26/2021         EPPlus Software AB       Modified to work with dxf styling for tables
 *************************************************************************************************/
using OfficeOpenXml.Constants;
using System;
using System.Xml;
namespace OfficeOpenXml.Style.Dxf
{
    /// <summary>
    /// Differential formatting record used in conditional formatting
    /// </summary>
    public class ExcelDxfStyle : ExcelDxfStyleBase
    {
        internal ExcelDxfStyle(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback)
            : base(nameSpaceManager, topNode, styles, callback)
        {
            NumberFormat = new ExcelDxfNumberFormat(_styles, callback);
            Font = new ExcelDxfFont(_styles, callback);
            Alignment = new ExcelDxfAlignment(styles, callback);
            Protection = new ExcelDxfProtection(styles, callback);

            if (topNode != null)
            {                
                NumberFormat.SetValuesFromXml(_helper);
                Font.SetValuesFromXml(_helper);
                Alignment.SetValuesFromXml(_helper);
                Protection.SetValuesFromXml(_helper);
            }
         }
        /// <summary>
        /// Font formatting settings
        /// </summary>
        public ExcelDxfFont Font { get; internal set; }
        /// <summary>
        /// Number format settings
        /// </summary>
        public ExcelDxfNumberFormat NumberFormat { get; internal set; }
        /// <summary>
        /// Cell alignment properties
        /// </summary>
        public ExcelDxfAlignment Alignment
        {
            get;
            internal set;
        }
        /// <summary>
        /// Cell protection properties used when the worksheet is protected.d
        /// </summary>
        public ExcelDxfProtection Protection
        {
            get;
            internal set;
        }
        bool _checkbox = false;
        /// <summary>
        /// If the cells applying this format should have a checkbox if the value of the cell is a boolean.
        /// </summary>
        public bool Checkbox
        {
            get
            {
                return _checkbox;
            }
            set
            {
                _checkbox = value;
                _callback(eStyleClass.Style, eStyleProperty.Checkbox, value);
            }
        }
        internal override string Id
        {
            get
            {
                return base.GetId() + Font.Id + NumberFormat.Id + Alignment.Id + Protection.Id + (Checkbox ? "1" : "0"); 
            }
        }
        /// <summary>
        /// If the object has any properties set
        /// </summary>
        public override bool HasValue
        {
            get
            {
                return base.HasValue || Font.HasValue || NumberFormat.HasValue || Alignment.HasValue || Protection.HasValue || Checkbox;
            }
        }
        internal override DxfStyleBase Clone()
        {
            var s = new ExcelDxfStyle(_helper.NameSpaceManager, null, _styles, _callback)
            {
                Font = (ExcelDxfFont)Font.Clone(),
                Fill = (ExcelDxfFill)Fill.Clone(),
                Border = (ExcelDxfBorderBase)Border.Clone(),
                NumberFormat = (ExcelDxfNumberFormat)NumberFormat.Clone(),
                Alignment = (ExcelDxfAlignment)Alignment.Clone(),
                Protection = (ExcelDxfProtection)Protection.Clone(),
                Checkbox = Checkbox
            };

            return s;
        }
        internal override void CreateNodes(XmlHelper helper, string path)
        {
            if (Font.HasValue) Font.CreateNodes(helper, "d:font");
            if (NumberFormat.HasValue) NumberFormat.CreateNodes(helper, "d:numFmt");
            if (Fill.HasValue) Fill.CreateNodes(helper, "d:fill");
            if (Alignment.HasValue) Alignment.CreateNodes(helper, "d:alignment");
            if (Border.HasValue) Border.CreateNodes(helper, "d:border");
            if (Protection.HasValue) Protection.CreateNodes(helper, "d:protection");
            if(Checkbox)
            {
                var cmplNode = (XmlElement)helper.CreateNode("d:extLst/d:ext/xfpb:DXFComplement");
                var extNode = (XmlElement)cmplNode.ParentNode;
                extNode.SetAttribute("xmlns:xfpb", Schemas.schemaFeaturePropertyBag);
                extNode.SetAttribute("uri", ExtLstUris.FeaturePropertyBag);
                cmplNode.SetAttribute("i", "0");
                _styles._hasDxfCheckbox = true;
            }
        }
        internal override void SetStyle()
        {
            if (_callback != null)
            {
                NumberFormat.SetStyle();
                base.SetStyle();
                Font.SetStyle();
                Alignment.SetStyle();
                Protection.SetStyle();
                _callback.Invoke(eStyleClass.Style, eStyleProperty.Checkbox, Checkbox);
            }
        }
        /// <summary>
        /// Clears all properties
        /// </summary>
        public override void Clear()
        {            
            base.Clear();
            Font.Clear();
            NumberFormat.Clear();
            Alignment.Clear();
            Protection.Clear();
            Checkbox = false;
        }
    }
}
