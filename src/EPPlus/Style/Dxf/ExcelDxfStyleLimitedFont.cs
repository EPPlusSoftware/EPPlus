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
using System.Xml;
namespace OfficeOpenXml.Style.Dxf
{
    /// <summary>
    /// Differential formatting record used in conditional formatting
    /// </summary>
    public class ExcelDxfStyleLimitedFont : ExcelDxfStyleBase
    {
        internal ExcelDxfStyleLimitedFont(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles, string dxfIdPath)
            : base(nameSpaceManager, topNode, styles, dxfIdPath)
        {
            Font = new ExcelDxfFont(styles);
            if (topNode != null)
            {
                Font.GetValuesFromXml(_helper);
            }
        }
        /// <summary>
        /// Font formatting settings
        /// </summary>
        public ExcelDxfFontBase Font { get; set; }

        /// <summary>
        /// Clone the object
        /// </summary>
        /// <returns>A new instance of the object</returns>
        protected internal override DxfStyleBase Clone()
        {
            var s = new ExcelDxfStyleLimitedFont(_helper.NameSpaceManager, null, _styles, _dxfIdPath)
            {
                Font = (ExcelDxfFont)Font.Clone(),
                NumberFormat = (ExcelDxfNumberFormat)NumberFormat.Clone(),
                Fill = (ExcelDxfFill)Fill.Clone(),
                Border = (ExcelDxfBorderBase)Border.Clone()
            };

            return s;
        }
        protected internal override void CreateNodes(XmlHelper helper, string path)
        {
            if (Font.HasValue) Font.CreateNodes(helper, "d:font");
            base.CreateNodes(helper, path);
        }
        public override bool HasValue
        {
            get
            {
                return Font.HasValue || base.HasValue;
            }
        }
        public override void Clear()
        {
            base.Clear();
            Font.Clear();
        }
    }
}
