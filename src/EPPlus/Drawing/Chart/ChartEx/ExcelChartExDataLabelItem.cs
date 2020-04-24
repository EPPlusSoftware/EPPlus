using OfficeOpenXml.Utils.Extentions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    public class ExcelChartExDataLabelItem : ExcelChartExDataLabel
    {
        internal ExcelChartExDataLabelItem(ExcelChart chart, XmlNamespaceManager nsm, XmlNode node) : base(chart, nsm, node)
        {

        }
    }
}
