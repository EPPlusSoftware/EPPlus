using OfficeOpenXml.Utils.Extentions;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// 
    /// </summary>
    public class ExcelChartExTitle : ExcelChartTitle
    {
        public ExcelChartExTitle(ExcelChart chart, XmlNamespaceManager nsm, XmlNode node) : base(chart, nsm, node, "cx")
        {
            
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
    }
}
