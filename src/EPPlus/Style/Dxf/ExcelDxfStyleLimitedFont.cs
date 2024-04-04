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
 *************************************************************************************************/
using System;
using System.Xml;
namespace OfficeOpenXml.Style.Dxf
{
    /// <summary>
    /// Differential formatting record with limited font settings
    /// </summary>
    public class ExcelDxfStyleLimitedFont : ExcelDxfStyleBase
    {
        internal ExcelDxfStyleLimitedFont(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback)
            : base(nameSpaceManager, topNode, styles, callback)
        {
            Font = new ExcelDxfFont(styles, _callback);
            if (topNode != null)
            {
                Font.SetValuesFromXml(_helper);
            }
        }
        /// <summary>
        /// Font formatting settings
        /// </summary>
        public ExcelDxfFontBase Font { get; internal set; }

        /// <summary>
        /// Clone the object
        /// </summary>
        /// <returns>A new instance of the object</returns>
        internal override DxfStyleBase Clone()
        {
            var s = new ExcelDxfStyleLimitedFont(_helper.NameSpaceManager, null, _styles, _callback)
            {
                Font = (ExcelDxfFont)Font.Clone(),
                Fill = (ExcelDxfFill)Fill.Clone(),
                Border = (ExcelDxfBorderBase)Border.Clone()
            };

            return s;
        }
        internal override void CreateNodes(XmlHelper helper, string path)
        {
            if (Font.HasValue) Font.CreateNodes(helper, "d:font");
            if (Fill.HasValue) Fill.CreateNodes(helper, "d:fill");
            if (Border.HasValue) Border.CreateNodes(helper, "d:border");
        }
        /// <summary>
        /// If the object has any properties set
        /// </summary>
        public override bool HasValue
        {
            get
            {
                return Font.HasValue || base.HasValue;
            }
        }
        internal override void SetStyle()
        {
            if (_callback != null)
            {
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
        }
        internal override string Id
        {
            get
            {
                return GetId() + ExcelDxfNumberFormat.GetEmptyId() + ExcelDxfAlignment.GetEmptyId() + ExcelDxfProtection.GetEmptyId();                    
            }
        }
		internal override string GetId()
		{
			return base.GetId() + Font.Id;
		}
	}
}
