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
using System.Globalization;
using System.Linq;
using System.Xml;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart
{
    public abstract class ExcelStandardChartWithLines : ExcelChartStandard, IDrawingDataLabel
    {
        internal ExcelStandardChartWithLines(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent = null) :
            base(drawings, node, uriChart, part, chartXml, chartNode, parent)
        {
        }

        internal ExcelStandardChartWithLines(ExcelChart topChart, XmlNode chartNode, ExcelGroupShape parent = null) :
            base(topChart, chartNode, parent)
        {
        }
        internal ExcelStandardChartWithLines(ExcelDrawings drawings, XmlNode node, eChartType? type, ExcelChart topChart, ExcelPivotTable PivotTableSource, XmlDocument chartXml, ExcelGroupShape parent = null) :
            base(drawings, node, type, topChart, PivotTableSource, chartXml, parent)
        {

        }
        string MARKER_PATH = "c:marker/@val";
        /// <summary>
        /// If the series has markers
        /// </summary>
        public bool Marker
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeBool(MARKER_PATH, false);
            }
            set
            {
                _chartXmlHelper.SetXmlNodeBool(MARKER_PATH, value, false);
            }
        }

        string SMOOTH_PATH = "c:smooth/@val";
        /// <summary>
        /// If the series has smooth lines
        /// </summary>
        public bool Smooth
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeBool(SMOOTH_PATH, false);
            }
            set
            {
                if (ChartType == eChartType.Line3D)
                {
                    throw new ArgumentException("Smooth", "Smooth does not apply to 3d line charts");
                }
                _chartXmlHelper.SetXmlNodeBool(SMOOTH_PATH, value);
            }
        }
        //string _chartTopPath = "c:chartSpace/c:chart/c:plotArea/{0}";
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
                    _dataLabel = new ExcelChartDataLabelStandard(this, NameSpaceManager, ChartNode, "dLbls", _chartXmlHelper.SchemaNodeOrder);
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
        const string _gapWidthPath = "c:upDownBars/c:gapWidth/@val";
        /// <summary>
        /// The gap width between the up and down bars
        /// </summary>
        public double? UpDownBarGapWidth
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeIntNull(_gapWidthPath);
            }
            set
            {
                if (value == null)
                {
                    _chartXmlHelper.DeleteNode(_gapWidthPath, true);
                }
                if (value < 0 || value > 500)
                {
                    throw (new ArgumentOutOfRangeException("GapWidth ranges between 0 and 500"));
                }
                _chartXmlHelper.SetXmlNodeString(_gapWidthPath, value.Value.ToString(CultureInfo.InvariantCulture));
            }
        }
        ExcelChartStyleItem _upBar = null;
        const string _upBarPath = "c:upDownBars/c:upBars";
        /// <summary>
        /// Format the up bars on the chart
        /// </summary>
        public ExcelChartStyleItem UpBar
        {
            get
            {
                return _upBar;
            }
        }
        ExcelChartStyleItem _downBar = null;
        const string _downBarPath = "c:upDownBars/c:downBars";
        /// <summary>
        /// Format the down bars on the chart
        /// </summary>
        public ExcelChartStyleItem DownBar
        {
            get
            {
                return _downBar;
            }
        }
        ExcelChartStyleItem _hiLowLines = null;
        const string _hiLowLinesPath = "c:hiLowLines";
        /// <summary>
        /// Format the high-low lines for the series.
        /// </summary>
        public ExcelChartStyleItem HighLowLine
        {
            get
            {
                return _hiLowLines;
            }
        }
        ExcelChartStyleItem _dropLines = null;
        const string _dropLinesPath = "c:dropLines";

        /// <summary>
        /// Format the drop lines for the series.
        /// </summary>
        public ExcelChartStyleItem DropLine
        {
            get
            {
                return _dropLines;
            }
        }
        /// <summary>
        /// Adds up and/or down bars to the chart.        
        /// </summary>
        /// <param name="upBars">Adds up bars if up bars does not exist.</param>
        /// <param name="downBars">Adds down bars if down bars does not exist.</param>
        public void AddUpDownBars(bool upBars = true, bool downBars = true)
        {
            if (upBars && _upBar == null)
            {
                _upBar = new ExcelChartStyleItem(NameSpaceManager, ChartNode, this, _upBarPath, RemoveUpBar);
                var chart = _topChart ?? this;
                chart.ApplyStyleOnPart(_upBar, chart.StyleManager?.Style?.UpBar);
            }
            if (downBars && _downBar == null)
            {
                _downBar = new ExcelChartStyleItem(NameSpaceManager, ChartNode, this, _downBarPath, RemoveDownBar);
                var chart = _topChart ?? this;
                chart.ApplyStyleOnPart(_upBar, chart.StyleManager?.Style?.DownBar);
            }
        }
        /// <summary>
        /// Adds droplines to the chart.        
        /// </summary>
        public ExcelChartStyleItem AddDropLines()
        {
            if (_dropLines == null)
            {
                _dropLines = new ExcelChartStyleItem(NameSpaceManager, ChartNode, this, _dropLinesPath, RemoveDropLines);
                var chart = _topChart ?? this;
                chart.ApplyStyleOnPart(_upBar, chart.StyleManager?.Style?.DropLine);
            }
            return _dropLines;
        }
        /// <summary>
        /// Adds High-Low lines to the chart.        
        /// </summary>
        public ExcelChartStyleItem AddHighLowLines()
        {
            if (_hiLowLines == null)
            {
                _hiLowLines = new ExcelChartStyleItem(NameSpaceManager, ChartNode, this, _hiLowLinesPath, RemoveHiLowLines);
                var chart = _topChart ?? this;
                chart.ApplyStyleOnPart(_upBar, chart.StyleManager?.Style?.HighLowLine);
            }
            return HighLowLine;
        }
        //TODO: Consider adding this method later (for all charts with datalabels)
        ///// <summary>
        ///// Adds datalabels to the chart
        ///// </summary>
        ///// <param name="position">The position of the datalabels</param>
        ///// <returns></returns>
        //public ExcelChartDataLabel AddDataLabels(eLabelPosition position=eLabelPosition.Center)
        //{
        //    DataLabel.Position = position;
        //    var chart = _topChart ?? this;
        //    foreach (var serie in chart.Series)
        //    {
        //        if (serie is IDrawingSerieDataLabel dl)
        //            dl.DataLabel.Position = position;
        //        if (chart.StyleManager.StylePart != null)
        //        {
        //            chart.StyleManager.ApplyStyle(serie, chart.StyleManager.Style.DataLabel);
        //        }

        //    }
        //    return DataLabel;
        //}
        internal override eChartType GetChartType(string name)
        {
            if (name == "lineChart")
            {
                if (Marker)
                {
                    if (Grouping == eGrouping.Stacked)
                    {
                        return eChartType.LineMarkersStacked;
                    }
                    else if (Grouping == eGrouping.PercentStacked)
                    {
                        return eChartType.LineMarkersStacked100;
                    }
                    else
                    {
                        return eChartType.LineMarkers;
                    }
                }
                else
                {
                    if (Grouping == eGrouping.Stacked)
                    {
                        return eChartType.LineStacked;
                    }
                    else if (Grouping == eGrouping.PercentStacked)
                    {
                        return eChartType.LineStacked100;
                    }
                    else
                    {
                        return eChartType.Line;
                    }
                }
            }
            else if (name == "line3DChart")
            {
                return eChartType.Line3D;
            }
            return base.GetChartType(name);
        }
        internal override void InitSeries(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot, List<ExcelChartSerie> list = null)
        {
            base.InitSeries(chart, ns, node, isPivot, list);
            AddSchemaNodeOrder(SchemaNodeOrder, new string[] { "gapWidth", "upbars", "downbars" });
            Series.Init(chart, ns, node, isPivot, base.Series._list);

            //Up bars
            if (_upBar==null && ExistsNode(node, _upBarPath))
            {
                _upBar = new ExcelChartStyleItem(ns, node, this, _upBarPath, RemoveUpBar);
            }

            //Down bars
            if (_downBar == null && ExistsNode(node, _downBarPath))
            {
                _downBar = new ExcelChartStyleItem(ns, node, this, _downBarPath, RemoveDownBar);
            }

            //Drop lines
            if (_dropLines == null && ExistsNode(node, _dropLinesPath))
            {
                _dropLines = new ExcelChartStyleItem(ns, node, this, _dropLinesPath, RemoveDropLines);
            }

            //High / low lines
            if (_hiLowLines == null && ExistsNode(node, _hiLowLinesPath))
            {
                _hiLowLines = new ExcelChartStyleItem(ns, node, this, _hiLowLinesPath, RemoveHiLowLines);
            }


        }

        /// <summary>
        /// The series for the chart
        /// </summary>
        public new ExcelChartSeries<ExcelLineChartSerie> Series
        {
            get;
        } = new ExcelChartSeries<ExcelLineChartSerie>();
        #region Remove Line/Bar
        private void RemoveUpBar()
        {
            _upBar = null;
        }
        private void RemoveDownBar()
        {
            _downBar = null;
        }
        private void RemoveDropLines()
        {
            _dropLines = null;
        }
        private void RemoveHiLowLines()
        {
            _hiLowLines = null;
        }
        #endregion

    }

    /// <summary>
    /// Provides access to line chart specific properties
    /// </summary>
    public class ExcelLineChart : ExcelStandardChartWithLines
    {
        #region "Constructors"
        internal ExcelLineChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent = null) :
            base(drawings, node, uriChart, part, chartXml, chartNode, parent)
        {
        }

        internal ExcelLineChart (ExcelChart topChart, XmlNode chartNode, ExcelGroupShape parent = null) :
            base(topChart, chartNode, parent)
        {
        }
        internal ExcelLineChart(ExcelDrawings drawings, XmlNode node, eChartType? type, ExcelChart topChart, ExcelPivotTable PivotTableSource, XmlDocument chartXml, ExcelGroupShape parent = null) :
            base(drawings, node, type, topChart, PivotTableSource, chartXml, parent)
        {
            if (type != eChartType.Line3D)
            {
                Smooth = false;
            }
        }
        #endregion
    }
}
