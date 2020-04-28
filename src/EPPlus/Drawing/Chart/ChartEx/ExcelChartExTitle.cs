using OfficeOpenXml.Utils.Extentions;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    public class ExcelChartExTitle : ExcelChartTitle
    {
        public ExcelChartExTitle(ExcelChart chart, XmlNamespaceManager nsm, XmlNode node) : base(chart, nsm, node, "cx")
        {
            SchemaNodeOrder = new string[] { "tx", "spPr", "txPr" };
        }
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
