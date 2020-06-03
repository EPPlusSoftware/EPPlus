/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/25/2020         EPPlus Software AB       Added this class
 *************************************************************************************************/
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Provides access to stock chart specific properties
    /// </summary>
    public class ExcelStockChart : ExcelStandardChartWithLines
    {
        internal ExcelStockChart(ExcelDrawings drawings, XmlNode node, eChartType? type, ExcelChart topChart, ExcelPivotTable PivotTableSource, XmlDocument chartXml, ExcelGroupShape parent = null) :
            base(drawings, node, type, topChart, PivotTableSource, chartXml, parent)
        {
            if(type==eChartType.StockVHLC || type==eChartType.StockVOHLC)
            {
                var barChart = new ExcelBarChart(this, _chartNode.PreviousSibling, parent);
                barChart.Direction = eDirection.Column;
                _plotArea = new ExcelChartPlotArea(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:plotArea", NameSpaceManager), barChart, "c", this);
            }
        }

        internal ExcelStockChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent = null) :
            base(drawings, node, uriChart, part, chartXml, chartNode, parent)
        {
            if(chartNode.LocalName=="barChart")
            {
                var barChart = new ExcelBarChart(this, chartNode, parent);
                barChart.Direction = eDirection.Column;
                _plotArea = new ExcelChartPlotArea(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:plotArea", NameSpaceManager), barChart, "c", this);
                _chartNode = chartNode.NextSibling;
            }
        }
        internal ExcelStockChart(ExcelChart topChart, XmlNode chartNode, ExcelGroupShape parent = null) :
            base(topChart, chartNode, parent)
        {
        }
        internal override void InitSeries(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot, List<ExcelChartSerie> list = null)
        {
            base.InitSeries(chart, ns, node, isPivot, list);
            Series.Init(chart, ns, node, isPivot, base.Series._list);
        }
        /// <summary>
        /// A collection of series for a Stock Chart
        /// </summary>
        public new ExcelChartSeries<ExcelStockChartSerie> Series { get; } = new ExcelChartSeries<ExcelStockChartSerie>();
    }
}
