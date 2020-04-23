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
        internal void Init(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot, List<ExcelChartSerie> list = null)
        {
            _ns = ns;
            _chart = chart;
            _node = node;
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
            //SchemaNodeOrder = new string[] { "view3D", "plotArea", "barDir", "grouping", "scatterStyle", "varyColors", "ser", "marker", "invertIfNegative", "pictureOptions", "dPt", "explosion", "dLbls", "firstSliceAng", "holeSize", "shape", "legend", "axId" };
            foreach (XmlNode n in node.SelectNodes("c:ser", ns))
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
                    case "stockChart":
                        s = new ExcelLineChartSerie(_chart, ns, n, isPivot);
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
                        s = new ExcelChartSerie(_chart, ns, n, isPivot);
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
                return _list.Count;
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
            return AddSeries(Serie.FullAddressAbsolute, XSerie.FullAddressAbsolute, "");
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
            if(_list.Count==256)
            {
                throw (new InvalidOperationException("Charts have a maximum of 256 series."));
            }
            XmlElement ser = _node.OwnerDocument.CreateElement("c", "ser", ExcelPackage.schemaChart);
            XmlNodeList node = _node.SelectNodes("c:ser", _ns);
            if (node.Count > 0)
            {
                _node.InsertAfter(ser, node[node.Count - 1]);
            }
            else
            {
                var f = XmlHelperFactory.Create(_ns, _node);
                f.InserAfter(_node, "c:varyColors,c:grouping,c:barDir,c:scatterStyle,c:ofPieType", ser);
            }

            //If the chart is added from a chart template, then use the chart templates series xml
            if (!string.IsNullOrEmpty(_chart._drawings._seriesTemplateXml))
            {
                ser.InnerXml = _chart._drawings._seriesTemplateXml;
            }

            int idx = FindIndex();
            ser.InnerXml = string.Format("<c:idx val=\"{1}\" /><c:order val=\"{1}\" /><c:tx><c:strRef><c:f></c:f><c:strCache><c:ptCount val=\"1\" /></c:strCache></c:strRef></c:tx>{2}{5}{0}{3}{4}", AddExplosion(Chart.ChartType), idx, AddSpPrAndScatterPoint(Chart.ChartType), AddAxisNodes(Chart.ChartType), AddSmooth(Chart.ChartType), AddMarker(Chart.ChartType));
            ExcelChartSerie serie;
            switch (Chart.ChartType)
            {
                case eChartType.Bubble:
                case eChartType.Bubble3DEffect:
                    serie = new ExcelBubbleChartSerie(_chart, _ns, ser, _isPivot)
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
                    serie = new ExcelScatterChartSerie(_chart, _ns, ser, _isPivot);
                    break;
                case eChartType.Radar:
                case eChartType.RadarFilled:
                case eChartType.RadarMarkers:
                    serie = new ExcelRadarChartSerie(_chart, _ns, ser, _isPivot);
                    break;
                case eChartType.Surface:
                case eChartType.SurfaceTopView:
                case eChartType.SurfaceTopViewWireframe:
                case eChartType.SurfaceWireframe:
                    serie = new ExcelSurfaceChartSerie(_chart, _ns, ser, _isPivot);
                    break;
                case eChartType.Pie:
                case eChartType.Pie3D:
                case eChartType.PieExploded:
                case eChartType.PieExploded3D:
                case eChartType.PieOfPie:
                case eChartType.Doughnut:
                case eChartType.DoughnutExploded:
                case eChartType.BarOfPie:
                    serie = new ExcelPieChartSerie(_chart, _ns, ser, _isPivot);
                    break;
                case eChartType.Line:
                case eChartType.LineMarkers:
                case eChartType.LineMarkersStacked:
                case eChartType.LineMarkersStacked100:
                case eChartType.LineStacked:
                case eChartType.LineStacked100:
                case eChartType.Line3D:
                    serie = new ExcelLineChartSerie(_chart, _ns, ser, _isPivot);
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
                    serie = new ExcelBarChartSerie(_chart, _ns, ser, _isPivot);
                    ((ExcelBarChartSerie)serie).InvertIfNegative = false;
                    break;
                case eChartType.Area:
                case eChartType.Area3D:
                case eChartType.AreaStacked:
                case eChartType.AreaStacked100:
                case eChartType.AreaStacked1003D:
                case eChartType.AreaStacked3D:
                    serie = new ExcelAreaChartSerie(_chart, _ns, ser, _isPivot);
                    break;
                default:
                    serie = new ExcelChartSerie(_chart, _ns, ser, _isPivot);
                    break;
            }
            serie.Series = SerieAddress;
            serie.XSeries = XSerieAddress;
            _list.Add((T)serie);
            if (_chart.StyleManager.StylePart != null)
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
        private int FindIndex()
        {
            int ret = 0, newID = 0;
            if (_chart.PlotArea.ChartTypes.Count > 1)
            {
                foreach (var chart in _chart.PlotArea.ChartTypes)
                {
                    if (newID > 0)
                    {
                        foreach (ExcelChartSerie serie in chart.Series)
                        {
                            serie.SetID((++newID).ToString());
                        }
                    }
                    else
                    {
                        if (chart == _chart)
                        {
                            ret += _list.Count + 1;
                            newID = ret;
                        }
                        else
                        {
                            ret += chart.Series.Count;
                        }
                    }
                }
                return ret - 1;
            }
            else
            {
                return _list.Count;
            }
        }
        #endregion
        #region "Xml init Functions"
        private string AddMarker(eChartType chartType)
        {
            if (chartType == eChartType.Line ||
                chartType == eChartType.LineStacked ||
                chartType == eChartType.LineStacked100 ||
                chartType == eChartType.XYScatterLines ||
                chartType == eChartType.XYScatterSmooth ||
                chartType == eChartType.XYScatterLinesNoMarkers ||
                chartType == eChartType.XYScatterSmoothNoMarkers)
            {
                return "<c:marker><c:symbol val=\"none\" /></c:marker>";
            }
            else
            {
                return "";
            }
        }
        private string AddSpPrAndScatterPoint(eChartType chartType)
        {
            if (chartType == eChartType.XYScatter)
            {
                return "<c:spPr><a:noFill/><a:ln w=\"28575\"><a:noFill /></a:ln><a:effectLst/><a:sp3d/></c:spPr>";
            }
            else
            {
                return "";
            }
        }
        private string AddAxisNodes(eChartType chartType)
        {
            if (chartType == eChartType.XYScatter ||
                 chartType == eChartType.XYScatterLines ||
                 chartType == eChartType.XYScatterLinesNoMarkers ||
                 chartType == eChartType.XYScatterSmooth ||
                 chartType == eChartType.XYScatterSmoothNoMarkers ||
                 chartType == eChartType.Bubble ||
                 chartType == eChartType.Bubble3DEffect)
            {
                return "<c:xVal /><c:yVal />";
            }
            else
            {
                return "<c:val />";
            }
        }

        private string AddExplosion(eChartType chartType)
        {
            if (chartType == eChartType.PieExploded3D ||
               chartType == eChartType.PieExploded ||
                chartType == eChartType.DoughnutExploded)
            {
                return "<c:explosion val=\"25\" />"; //Default 25;
            }
            else
            {
                return "";
            }
        }
        private string AddSmooth(eChartType chartType)
        {
            if (chartType == eChartType.XYScatterSmooth ||
               chartType == eChartType.XYScatterSmoothNoMarkers)
            {
                return "<c:smooth val=\"1\" />"; //Default 25;
            }
            else
            {
                return "";
            }
        }
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
        #endregion
    }
}
