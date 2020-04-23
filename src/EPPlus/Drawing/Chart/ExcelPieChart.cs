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
using System.Collections.Generic;
using System.Xml;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Provides access to pie chart specific properties
    /// </summary>
    public class ExcelPieChart : ExcelChartStandard, IDrawingDataLabel
    {
        internal ExcelPieChart(ExcelDrawings drawings, XmlNode node, eChartType type, bool isPivot, ExcelGroupShape parent = null) :
            base(drawings, node, type, isPivot, parent)
        {
            
        }
        internal ExcelPieChart(ExcelDrawings drawings, XmlNode node, eChartType? type, ExcelChart topChart, ExcelPivotTable PivotTableSource, XmlDocument chartXml, ExcelGroupShape parent = null) :
            base(drawings, node, type, topChart, PivotTableSource, chartXml, parent)
        {
            
        }

        internal ExcelPieChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent = null) :
           base(drawings, node, uriChart, part, chartXml, chartNode, parent)
        {
            
        }
        internal ExcelPieChart(ExcelChart topChart, XmlNode chartNode, ExcelGroupShape parent = null) :
            base(topChart, chartNode, parent)
        {
            
        }
        internal override void InitSeries(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot, List<ExcelChartSerie> list = null)
        {
            base.InitSeries(chart, ns, node, isPivot, list);
            Series.Init(chart, ns, node, isPivot, base.Series._list);
        }

        ExcelChartDataLabel _dataLabel = null;
        /// <summary>
        /// Access to datalabel properties
        /// </summary>
        public ExcelChartDataLabel DataLabel
        {
            get
            {
                if (_dataLabel == null)
                {
                    _dataLabel = new ExcelChartDataLabel(this, NameSpaceManager, ChartNode, "dLbls", _chartXmlHelper.SchemaNodeOrder);
                }
                return _dataLabel;
            }
        }
        /// <summary>
        /// If the chart has datalabel
        /// </summary>
        public bool HasDataLabel
        {
            get
            {
                return ChartNode.SelectSingleNode("c:dLbls", NameSpaceManager) != null;
            }
        }
        internal override eChartType GetChartType(string name)
        {
            if (name == "pieChart")
            {
                if (Series.Count > 0 && (Series[0]).Explosion>0)
                {
                    return eChartType.PieExploded;
                }
                else
                {
                    return eChartType.Pie;
                }
            }
            else if (name == "pie3DChart")
            {
                if (Series.Count > 0 && (Series[0]).Explosion > 0)
                {
                    return eChartType.PieExploded3D;
                }
                else
                {
                    return eChartType.Pie3D;
                }
            }
            return base.GetChartType(name);
        }
        /// <summary>
        /// A collection of series for a Pie Chart
        /// </summary>
        public new ExcelChartSeries<ExcelPieChartSerie> Series { get; } = new ExcelChartSeries<ExcelPieChartSerie>();

    }
}
