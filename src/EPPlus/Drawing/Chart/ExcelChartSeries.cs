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
using System.Collections;
using OfficeOpenXml.Table.PivotTable;
using System.Linq;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Utils;
namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Collection class for chart series
    /// </summary>
    public class ExcelChartSeries<T> : IEnumerable<T> where T : ExcelChartSerie
    {
        internal List<ExcelChartSerie> _list;
        internal ExcelChart _chart;
        XmlNode _node;
        XmlNamespaceManager _ns;
        internal void Init(ExcelChart chart, XmlNamespaceManager ns, XmlNode chartNode, bool isPivot, List<ExcelChartSerie> list = null)
        {
            _ns = ns;
            _chart = chart;
            _node = chartNode;
            _isPivot = isPivot;
            if (list == null)
            {
                _list = new List<ExcelChartSerie>();
            }
            else
            {
                _list = list;
                return;
            }

            if (_chart._isChartEx)
            {
                AddSeriesChartEx((ExcelChartEx)chart, ns, chartNode);
            }
            else
            {
                AddSeriesStandard(chart, ns, chartNode, isPivot);
            }
        }
        private void AddSeriesChartEx(ExcelChartEx chart, XmlNamespaceManager ns, XmlNode chartNode)
        {
            var histoGramSeries = new List<XmlElement>();
            int index = 0;
            foreach (XmlElement serieElement in chartNode.SelectNodes("cx:plotArea/cx:plotAreaRegion/cx:series", ns))
            {
                switch (chart.ChartType)
                {
                    case eChartType.Treemap:
                        _list.Add(new ExcelTreemapChartSerie(chart, ns, serieElement));
                        break;
                    case eChartType.BoxWhisker:
                        _list.Add(new ExcelBoxWhiskerChartSerie(chart, ns, serieElement));
                        break;
                    case eChartType.Histogram:
                    case eChartType.Pareto:
                        if(serieElement.GetAttribute("layoutId") == "paretoLine")
                        {
                            histoGramSeries.Add(serieElement);
                        }
                        else
                        {
                            _list.Add(new ExcelHistogramChartSerie(chart, ns, serieElement, index));
                        }
                        break;
                    case eChartType.RegionMap:
                        _list.Add(new ExcelRegionMapChartSerie(chart, ns, serieElement));
                        break;
                    case eChartType.Waterfall:
                        _list.Add(new ExcelWaterfallChartSerie(chart, ns, serieElement));
                        break;
                    default:
                        _list.Add(new ExcelChartExSerie(chart, ns, serieElement));
                        break;
                }
                index++;
            }
            if (chart.ChartType == eChartType.Pareto)
            {
                foreach (var e in histoGramSeries)
                {
                    if (e.GetAttribute("layoutId") == "paretoLine")
                    {
                        if (ConvertUtil.TryParseIntString(e.GetAttribute("ownerIdx"), out int ownerId))
                        {
                            var serie=(ExcelHistogramChartSerie)_list.FirstOrDefault(x => ((ExcelHistogramChartSerie)x)._index == ownerId);
                            if(serie!=null)
                            {
                                serie.AddParetoLineFromSerie(e);
                            }
                        }
                    }
                }
            }
        }
        private void AddSeriesStandard(ExcelChart chart, XmlNamespaceManager ns, XmlNode chartNode, bool isPivot)
        {
            foreach (XmlNode n in chartNode.SelectNodes("c:ser", ns))
            {
                ExcelChartSerie s;
                switch (chart.ChartNode.LocalName)
                {
                    case "barChart":
                    case "bar3DChart":
                        s = new ExcelBarChartSerie(_chart, ns, n, isPivot);
                        break;
                    case "lineChart":
                    case "line3DChart":
                        s = new ExcelLineChartSerie(_chart, ns, n, isPivot);
                        break;
                    case "stockChart":
                        s = new ExcelStockChartSerie(_chart, ns, n, isPivot);
                        break;
                    case "scatterChart":
                        s = new ExcelScatterChartSerie(_chart, ns, n, isPivot);
                        break;
                    case "pieChart":
                    case "pie3DChart":
                    case "ofPieChart":
                    case "doughnutChart":
                        s = new ExcelPieChartSerie(_chart, ns, n, isPivot);
                        break;
                    case "bubbleChart":
                        s = new ExcelBubbleChartSerie(_chart, ns, n, isPivot);
                        break;
                    case "radarChart":
                        s = new ExcelRadarChartSerie(_chart, ns, n, isPivot);
                        break;
                    case "surfaceChart":
                    case "surface3DChart":
                        s = new ExcelSurfaceChartSerie(_chart, ns, n, isPivot);
                        break;
                    case "areaChart":
                    case "area3DChart":
                        s = new ExcelAreaChartSerie(_chart, ns, n, isPivot);
                        break;
                    default:
                        s = new ExcelChartStandardSerie(_chart, ns, n, isPivot);
                        break;
                }
                _list.Add((T)s);
            }
        }

        /// <summary>
        /// Returns the serie at the specified position.  
        /// </summary>
        /// <param name="PositionID">The position of the series.</param>
        /// <returns></returns>
        public T this[int PositionID]
        {
            get
            {
                return (T)(_list[PositionID]);
            }
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _list?.Count ?? 0;
            }
        }
        /// <summary>
        /// Delete the chart at the specific position
        /// </summary>
        /// <param name="PositionID">Zero based</param>
        public void Delete(int PositionID)
        {
            ExcelChartSerie ser = _list[PositionID];
            ser.TopNode.ParentNode.RemoveChild(ser.TopNode);
            _list.RemoveAt(PositionID);
        }
        /// <summary>
        /// A reference to the chart object
        /// </summary>
        public ExcelChart Chart
        {
            get
            {
                return _chart;
            }
        }
        #region "Add Series"

        /// <summary>
        /// Add a new serie to the chart. Do not apply to pivotcharts.
        /// </summary>
        /// <param name="Serie">The Y-Axis range</param>
        /// <param name="XSerie">The X-Axis range</param>
        /// <returns>The serie</returns>
        public virtual T Add(ExcelRangeBase Serie, ExcelRangeBase XSerie)
        {
            if (_chart.PivotTableSource != null)
            {
                throw (new InvalidOperationException("Can't add a serie to a pivotchart"));
            }
            return AddSeries(Serie.FullAddressAbsolute, XSerie?.FullAddressAbsolute, "");
        }
        /// <summary>
        /// Add a new serie to the chart.Do not apply to pivotcharts.
        /// </summary>
        /// <param name="SerieAddress">The Y-Axis range</param>
        /// <param name="XSerieAddress">The X-Axis range</param>
        /// <returns>The serie</returns>
        public virtual T Add(string SerieAddress, string XSerieAddress)
        {
            if (_chart.PivotTableSource != null)
            {
                throw (new InvalidOperationException("Can't add a serie to a pivotchart"));
            }
            return AddSeries(SerieAddress, XSerieAddress, "");
        }
        /// <summary>
        /// Adds a new serie to the chart
        /// </summary>
        /// <param name="SerieAddress">The Y-Axis range</param>
        /// <param name="XSerieAddress">The X-Axis range</param>
        /// <param name="bubbleSizeAddress">Bubble chart size</param>
        /// <returns></returns>
        internal protected T AddSeries(string SerieAddress, string XSerieAddress, string bubbleSizeAddress)
        {
            if (_list.Count == 256)
            {
                throw (new InvalidOperationException("Charts have a maximum of 256 series."));
            }
            XmlElement serElement;
            if (_chart._isChartEx)
            {
                serElement = ExcelChartExSerie.CreateSeriesAndDataElement((ExcelChartEx)_chart, !string.IsNullOrEmpty(XSerieAddress));
            }
            else
            {
                serElement = ExcelChartStandardSerie.CreateSerieElement(_chart);
            }
            ExcelChartSerie serie;
            switch (Chart.ChartType)
            {
                case eChartType.Bubble:
                case eChartType.Bubble3DEffect:
                    serie = new ExcelBubbleChartSerie(_chart, _ns, serElement, _isPivot)
                    {
                        Bubble3D = Chart.ChartType == eChartType.Bubble3DEffect,
                        Series = SerieAddress,
                        XSeries = XSerieAddress,
                        BubbleSize = bubbleSizeAddress
                    };
                    break;
                case eChartType.XYScatter:
                case eChartType.XYScatterLines:
                case eChartType.XYScatterLinesNoMarkers:
                case eChartType.XYScatterSmooth:
                case eChartType.XYScatterSmoothNoMarkers:
                    serie = new ExcelScatterChartSerie(_chart, _ns, serElement, _isPivot);
                    break;
                case eChartType.Radar:
                case eChartType.RadarFilled:
                case eChartType.RadarMarkers:
                    serie = new ExcelRadarChartSerie(_chart, _ns, serElement, _isPivot);
                    break;
                case eChartType.Surface:
                case eChartType.SurfaceTopView:
                case eChartType.SurfaceTopViewWireframe:
                case eChartType.SurfaceWireframe:
                    serie = new ExcelSurfaceChartSerie(_chart, _ns, serElement, _isPivot);
                    break;
                case eChartType.Pie:
                case eChartType.Pie3D:
                case eChartType.PieExploded:
                case eChartType.PieExploded3D:
                case eChartType.PieOfPie:
                case eChartType.Doughnut:
                case eChartType.DoughnutExploded:
                case eChartType.BarOfPie:
                    serie = new ExcelPieChartSerie(_chart, _ns, serElement, _isPivot);
                    break;
                case eChartType.Line:
                case eChartType.LineMarkers:
                case eChartType.LineMarkersStacked:
                case eChartType.LineMarkersStacked100:
                case eChartType.LineStacked:
                case eChartType.LineStacked100:
                case eChartType.Line3D:
                    serie = new ExcelLineChartSerie(_chart, _ns, serElement, _isPivot);
                    if (Chart.ChartType == eChartType.LineMarkers ||
                        Chart.ChartType == eChartType.LineMarkersStacked ||
                        Chart.ChartType == eChartType.LineMarkersStacked100)
                    {
                        ((ExcelLineChartSerie)serie).Marker.Style = eMarkerStyle.Square;
                    }
                    ((ExcelLineChartSerie)serie).Smooth = ((ExcelLineChart)Chart).Smooth;
                    break;
                case eChartType.BarClustered:
                case eChartType.BarStacked:
                case eChartType.BarStacked100:
                case eChartType.ColumnClustered:
                case eChartType.ColumnStacked:
                case eChartType.ColumnStacked100:
                case eChartType.BarClustered3D:
                case eChartType.BarStacked3D:
                case eChartType.BarStacked1003D:
                case eChartType.Column3D:
                case eChartType.ColumnClustered3D:
                case eChartType.ColumnStacked3D:
                case eChartType.ColumnStacked1003D:
                case eChartType.ConeBarClustered:
                case eChartType.ConeBarStacked:
                case eChartType.ConeBarStacked100:
                case eChartType.ConeCol:
                case eChartType.ConeColClustered:
                case eChartType.ConeColStacked:
                case eChartType.ConeColStacked100:
                case eChartType.CylinderBarClustered:
                case eChartType.CylinderBarStacked:
                case eChartType.CylinderBarStacked100:
                case eChartType.CylinderCol:
                case eChartType.CylinderColClustered:
                case eChartType.CylinderColStacked:
                case eChartType.CylinderColStacked100:
                case eChartType.PyramidBarClustered:
                case eChartType.PyramidBarStacked:
                case eChartType.PyramidBarStacked100:
                case eChartType.PyramidCol:
                case eChartType.PyramidColClustered:
                case eChartType.PyramidColStacked:
                case eChartType.PyramidColStacked100:
                    serie = new ExcelBarChartSerie(_chart, _ns, serElement, _isPivot);
                    ((ExcelBarChartSerie)serie).InvertIfNegative = false;
                    break;
                case eChartType.Area:
                case eChartType.Area3D:
                case eChartType.AreaStacked:
                case eChartType.AreaStacked100:
                case eChartType.AreaStacked1003D:
                case eChartType.AreaStacked3D:
                    serie = new ExcelAreaChartSerie(_chart, _ns, serElement, _isPivot);
                    break;
                case eChartType.StockHLC:
                case eChartType.StockOHLC:
                case eChartType.StockVHLC:
                case eChartType.StockVOHLC:
                    serie = new ExcelStockChartSerie(_chart, _ns, serElement, _isPivot);
                    break;
                case eChartType.Treemap:
                    serie = new ExcelTreemapChartSerie((ExcelChartEx)_chart, _ns, serElement);
                    break;
                case eChartType.BoxWhisker:
                    serie = new ExcelBoxWhiskerChartSerie((ExcelChartEx)_chart, _ns, serElement);
                    break;
                case eChartType.Histogram:
                case eChartType.Pareto:
                    serie=new ExcelHistogramChartSerie((ExcelChartEx)_chart, _ns, serElement);
                    if(Chart.ChartType== eChartType.Pareto)
                    {
                        ((ExcelHistogramChartSerie)serie).AddParetoLine();
                    }
                    break;
                case eChartType.RegionMap:
                    serie = new ExcelRegionMapChartSerie((ExcelChartEx)_chart, _ns, serElement);
                    break;
                case eChartType.Waterfall:
                    serie = new ExcelWaterfallChartSerie((ExcelChartEx)_chart, _ns, serElement);
                    break;
                case eChartType.Sunburst:
                case eChartType.Funnel:
                    serie = new ExcelChartExSerie((ExcelChartEx)_chart, _ns, serElement);
                    break;
                default:
                    serie = new ExcelChartStandardSerie(_chart, _ns, serElement, _isPivot);
                    break;
            }
            serie.Series = SerieAddress;
            if (!string.IsNullOrEmpty(XSerieAddress))
            {
                serie.XSeries = XSerieAddress;
            }
            _list.Add((T)serie);
            if (_chart.StyleManager.StylePart != null && _chart._isChartEx == false)
            {
                _chart.StyleManager.ApplySeries();
            }
            return (T)serie;
        }
        bool _isPivot;
        internal void AddPivotSerie(ExcelPivotTable pivotTableSource)
        {
            var r = pivotTableSource.WorkSheet.Cells[pivotTableSource.Address.Address];
            _isPivot = true;
            AddSeries(r.Offset(0, 1, r._toRow - r._fromRow + 1, 1).FullAddressAbsolute, r.Offset(0, 0, r._toRow - r._fromRow + 1, 1).FullAddressAbsolute, "");
        }
        #endregion
        /// <summary>
        /// Gets the enumerator for the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<T> GetEnumerator()
        {
            return _list.Cast<T>().GetEnumerator();
        }
        IEnumerator<T> IEnumerable<T>.GetEnumerator()
        {
            return _list.Cast<T>().GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
    }
    public class ExcelHistogramChartSeries : ExcelChartSeries<ExcelHistogramChartSerie>
    {
        public void AddParetoLine()
        {
            if(_chart.ChartType==eChartType.Pareto)
            {
                return;
            }
            if (_chart.Axis.Length == 2)
            {
                //Add pareto axis
                var axis2 = (XmlElement)_chart._chartXmlHelper.CreateNode("cx:plotArea/cx:axis", false, true);
                axis2.SetAttribute("id", "2");
                axis2.InnerXml = "<cx:valScaling min=\"0\" max=\"1\"/><cx:units unit=\"percentage\"/><cx:tickLabels/>";
            }
            foreach(ExcelHistogramChartSerie ser in _list)
            {
                ser.AddParetoLineFromSerie((XmlElement)ser.TopNode);                
            }
        }
        public void RemoveParetoLine()
        {
            if (_chart.ChartType == eChartType.Histogram)
            {
                return;
            }
            if (_chart.Axis.Length == 2)
            {
                if (_chart.Axis.Length == 3)
                {
                    //Remove percentage axis
                    _chart.Axis[2].TopNode.ParentNode.RemoveChild(_chart.Axis[2].TopNode);
                    ((ExcelChartEx)_chart)._exAxis = null;
                    _chart._axis = new ExcelChartAxis[] { _chart._axis[0], _chart._axis[1] };
                }
            }
        }
    }
}
