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
using System.Text;
using System.Xml;
using OfficeOpenXml.Table.PivotTable;
using System.Globalization;
using OfficeOpenXml.Drawing.Interfaces;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Represents an Area Chart
    /// </summary>
    public sealed class ExcelAreaChart : ExcelChart, IDrawingDataLabel
    {
        #region "Constructors"
        internal ExcelAreaChart(ExcelDrawings drawings, XmlNode node, eChartType? type, ExcelChart topChart, ExcelPivotTable PivotTableSource, XmlDocument chartXml, ExcelGroupShape parent = null) :
            base(drawings, node, type, topChart, PivotTableSource, chartXml, parent)
        {
        }

        internal ExcelAreaChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent = null) :
           base(drawings, node, uriChart, part, chartXml, chartNode, parent)
        {
        }

        internal ExcelAreaChart(ExcelChart topChart, XmlNode chartNode, ExcelGroupShape parent = null) : 
            base(topChart, chartNode, parent)
        {
        }
        #endregion
        internal override void InitSeries(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot, List<ExcelChartSerie> list = null)
        {
            base.InitSeries(chart, ns, node, isPivot, list);
            Series.Init(chart, ns, node, isPivot, base.Series._list);
        }        
        #region DataLabel
        ExcelChartDataLabel _DataLabel = null;
        /// <summary>
        /// Access to datalabel properties
        /// </summary>
        public ExcelChartDataLabel DataLabel
        {
            get
            {
                if (_DataLabel == null)
                {
                    _DataLabel = new ExcelChartDataLabel(this, NameSpaceManager, ChartNode, "dLbls", _chartXmlHelper.SchemaNodeOrder);
                }
                return _DataLabel;
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
        #endregion
        /// <summary>
        /// The series for the Area Chart
        /// </summary>s
        public new ExcelChartSeries<ExcelAreaChartSerie> Series { get; } = new ExcelChartSeries<ExcelAreaChartSerie>();
    }
}
