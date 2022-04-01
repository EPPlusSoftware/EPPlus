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
            if (topNode != null)
            {                
                NumberFormat.SetValuesFromXml(_helper);
                Font.SetValuesFromXml(_helper);
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
        internal override string Id
        {
            get
            {
                return base.Id + Font.Id + NumberFormat.Id;
            }
        }
        /// <summary>
        /// If the object has any properties set
        /// </summary>
        public override bool HasValue
        {
            get
            {
                return base.HasValue || Font.HasValue || NumberFormat.HasValue;
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
            };

            return s;
        }
        internal override void CreateNodes(XmlHelper helper, string path)
        {
            if (Font.HasValue) Font.CreateNodes(helper, "d:font");
            if (NumberFormat.HasValue) NumberFormat.CreateNodes(helper, "d:numFmt");
            if (Fill.HasValue) Fill.CreateNodes(helper, "d:fill");
            if (Border.HasValue) Border.CreateNodes(helper, "d:border");
        }
        internal override void SetStyle()
        {
            if (_callback != null)
            {
                NumberFormat.SetStyle();
                base.SetStyle();
                Font.SetStyle();
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
        }
    }
}
