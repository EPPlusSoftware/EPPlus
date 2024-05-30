/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/29/2020         EPPlus Software AB           EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.Utils.Extensions;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// 
    /// </summary>
    public class ExcelChartExTitle : ExcelChartTitle
    {
        internal ExcelChartExTitle(ExcelChart chart, XmlNamespaceManager nsm, XmlNode node) : base(chart, nsm, node, "cx")
        {
            
        }
        /// <summary>
        /// Richtext.Text shorthand
        /// </summary>
        public override string Text
        {
            get
            {
                return RichText.Text;
            }
            set
            {
                var applyStyle = (RichText.Count == 0);
                RichText.Text = value;
                if (applyStyle) _chart.ApplyStyleOnPart(this, _chart.StyleManager?.Style?.Title, true);
            }

        }
        /// <summary>
        /// The side position alignment of the title
        /// </summary>
        public ePositionAlign PositionAlignment
        { 
            get
            {
                return GetXmlNodeString("@align").Replace("ctr", "center").ToEnum(ePositionAlign.Center);
            }
            set
            {
                SetXmlNodeString("@align", value.ToEnumString().Replace("center","ctr"));
            }
        }
        /// <summary>
        /// The position if the title
        /// </summary>
        public eSidePositions Position
        {
            get
            {
                switch(GetXmlNodeString("@pos"))
                {
                    case "l":
                        return eSidePositions.Left;
                    case "r":
                        return eSidePositions.Right;
                    case "b":
                        return eSidePositions.Bottom;
                    default:
                        return eSidePositions.Top;
                }
            }
            set
            {
                SetXmlNodeString("@align", value.ToEnumString().Substring(0,1).ToLowerInvariant());
            }
        }
        internal override bool HasLinkedCell => false;
    }
}
