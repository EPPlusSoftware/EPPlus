using OfficeOpenXml.Utils.Extentions;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    public class ExcelChartExLegend : ExcelChartLegend
    {
        internal ExcelChartExLegend(ExcelChartBase chart, XmlNamespaceManager nsm, XmlNode node) : base(nsm, node, chart, "cx")
        {

        }
        public ePositionAlign PositionAlignment
        {
            get
            {
                return GetXmlNodeString("@align").Replace("ctr", "center").ToEnum(ePositionAlign.Center);
            }
            set
            {
                SetXmlNodeString("@align", value.ToEnumString().Replace("center", "ctr"));
            }
        }
        /// <summary>
        /// The position of the Legend
        /// </summary>
        public override eLegendPosition Position
        {
            get
            {
                switch (GetXmlNodeString("@pos"))
                {
                    case "l":
                        return eLegendPosition.Left;
                    case "r":
                        return eLegendPosition.Right;
                    case "b":
                        return eLegendPosition.Bottom;
                    default:
                        return eLegendPosition.Top;
                }
            }
            set
            {
                if (value == eLegendPosition.TopRight)
                {
                    throw new InvalidOperationException("TopRight can not be set for Extended charts. Please use Top and set the PositionAlignment property.");
                }
                SetXmlNodeString("@align", value.ToEnumString().Substring(0, 1).ToLowerInvariant());
            }
        }
        /// <summary>
        /// Adds a legend to the chart
        /// </summary>
        public override void Add()
        {
            if (TopNode != null) return;

            //XmlHelper xml = new XmlHelper(NameSpaceManager, _chart.ChartXml);
            XmlHelper xml = XmlHelperFactory.Create(NameSpaceManager, _chart.ChartXml);
            xml.SchemaNodeOrder = _chart.SchemaNodeOrder;

            xml.CreateNode("cx:chartSpace/cx:chart/cx:legend");
            TopNode = _chart.ChartXml.SelectSingleNode("c:chartSpace/cx:chart/cx:legend", NameSpaceManager);
        }
    }
}
