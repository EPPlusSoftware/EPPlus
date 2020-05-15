/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/15/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Xml;
namespace OfficeOpenXml.Drawing.ChartEx
{
    public class ExcelTreemapChart : ExcelChartEx
    {
        internal ExcelTreemapChart(ExcelDrawings drawings, XmlNode node, ExcelGroupShape parent) :
            base(drawings, node, parent)
        {            
        }

        internal ExcelTreemapChart(ExcelDrawings drawings, XmlNode drawingsNode, eChartType? type, XmlDocument chartXml = null, ExcelGroupShape parent = null) :
            base(drawings, drawingsNode, type, chartXml, parent)
        {
            
        }
        internal ExcelTreemapChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent = null) :
            base(drawings, node, uriChart, part, chartXml, chartNode, parent)
        {
        }
        internal override void InitSeries(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot, List<ExcelChartSerie> list = null)
        {
            base.InitSeries(chart, ns, node, isPivot, list);
            Series.Init(chart, ns, node, isPivot, base.Series._list);
        }
        public new ExcelChartSeries<ExcelChartTreemapSerie> Series
        {
            get;
        } = new ExcelChartSeries<ExcelChartTreemapSerie>();
    }
}
