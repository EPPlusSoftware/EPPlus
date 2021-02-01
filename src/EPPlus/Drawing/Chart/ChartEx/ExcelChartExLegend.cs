using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// A legend for an Extended chart
    /// </summary>
    public class ExcelChartExLegend : ExcelChartLegend
    {
        internal ExcelChartExLegend(ExcelChart chart, XmlNamespaceManager nsm, XmlNode node) : base(nsm, node, chart, "cx")
        {
            SchemaNodeOrder = new string[] { "spPr","txPr" };
        }
        /// <summary>
        /// The side position alignment of the legend
        /// </summary>
        public ePositionAlign PositionAlignment
        {
            get
            {
                return GetXmlNodeString("@align").Replace("ctr", "center").ToEnum(ePositionAlign.Center);
            }
            set
            {
                if (TopNode == null) Add();
                SetXmlNodeString("@align", value.ToEnumString().Replace("center", "ctr"));
            }
        }
        /// <summary>
        /// The position of the Legend.
        /// </summary>
        /// <remarks>Setting the Position to TopRight will set the <see cref="Position"/> to Right and the <see cref="PositionAlignment" /> to Min</remarks>
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
                if (TopNode == null) Add();
                if (value == eLegendPosition.TopRight)
                {
                    PositionAlignment = ePositionAlign.Min;
                    value = eLegendPosition.Right;
                }
                SetXmlNodeString("@pos", value.ToEnumString().Substring(0, 1).ToLowerInvariant());
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

            TopNode = xml.CreateNode("cx:chartSpace/cx:chart/cx:legend");
        }
    }
}
