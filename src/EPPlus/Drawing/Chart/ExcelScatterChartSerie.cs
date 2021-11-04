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
using System.Text;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Drawing.Interfaces;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// A serie for a scatter chart
    /// </summary>
    public sealed class ExcelScatterChartSerie : ExcelChartSerieWithHorizontalErrorBars, IDrawingSerieDataLabel, IDrawingChartMarker
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="chart">The chart</param>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="node">Topnode</param>
        /// <param name="isPivot">Is pivotchart</param>
        internal ExcelScatterChartSerie(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot) :
            base(chart, ns, node, isPivot)
        {
            if (chart.ChartType == eChartType.XYScatterLines ||
                chart.ChartType == eChartType.XYScatterSmooth ||
                chart.ChartType == eChartType.XYScatter)
            {
                Marker.Style = eMarkerStyle.Square;
            }

            if (chart.ChartType == eChartType.XYScatterSmooth ||
                chart.ChartType == eChartType.XYScatterSmoothNoMarkers)
            {
                Smooth = 1;
            }
            else if (chart.ChartType == eChartType.XYScatterLines || chart.ChartType == eChartType.XYScatterLinesNoMarkers || chart.ChartType == eChartType.XYScatter)

            {
                Smooth = 0;
            }
        }

        ExcelChartSerieDataLabel _dataLabel = null;
        /// <summary>
        /// Data label properties
        /// </summary>
        public ExcelChartSerieDataLabel DataLabel
        {
            get
            {
                if (_dataLabel == null)
                {
                    _dataLabel = new ExcelChartSerieDataLabel(_chart, NameSpaceManager, TopNode, SchemaNodeOrder);
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
                return TopNode.SelectSingleNode("c:dLbls", NameSpaceManager) != null;
            }
        }
        const string smoothPath = "c:smooth/@val";
        /// <summary>
        /// Smooth for scattercharts
        /// </summary>
        public int Smooth
        {
            get
            {
                return GetXmlNodeInt(smoothPath);
            }
            internal set
            {
                SetXmlNodeString(smoothPath, value.ToString());
            }
        }
        const string markerPath = "c:marker/c:symbol/@val";
        ExcelChartMarker _chartMarker = null;
        /// <summary>
        /// A reference to marker properties
        /// </summary>
        public ExcelChartMarker Marker
        {
            get
            {
                if (IsMarkersAllowed() == false)
                {
                    return null;
                }

                if (_chartMarker == null)
                {
                    _chartMarker = new ExcelChartMarker(_chart, NameSpaceManager, TopNode, SchemaNodeOrder);
                }
                return _chartMarker;
            }
        }
        /// <summary>
        /// If the serie has markers
        /// </summary>
        /// <returns>True if serie has markers</returns>
        public bool HasMarker()
        {
            if (IsMarkersAllowed())
            {
                return ExistsNode("c:marker");
            }
            return false;
        }
        private bool IsMarkersAllowed()
        {
            var type = _chart.ChartType;
            if (type == eChartType.XYScatterLinesNoMarkers || type == eChartType.XYScatterSmoothNoMarkers)
            {
                return false;
            }
            return true;
        }
        ExcelChartDataPointCollection _dataPoints = null;
        /// <summary>
        /// A collection of the individual datapoints
        /// </summary>
        public ExcelChartDataPointCollection DataPoints
        {
            get
            {

                if (_dataPoints == null)
                {
                    _dataPoints = new ExcelChartDataPointCollection(_chart, NameSpaceManager, TopNode, SchemaNodeOrder);
                }
                return _dataPoints;
            }
        }        
        /// <summary>
        /// Line color.
        /// </summary>
        ///
        /// <value>
        /// The color of the line.
        /// </value>
        [Obsolete("Please use Border.Fill.Color property")]
        public Color LineColor
        {
            get
            {
                if (Border.Fill.Style == eFillStyle.SolidFill && Border.Fill.SolidFill.Color.ColorType == eDrawingColorType.Rgb)
                {
                    return Border.Fill.Color;
                }
                else
                {
                    return Color.Black;
                }
            }
            set
            {
                Border.Fill.Color = value;
            }
        }
        /// <summary>
        /// Gets or sets the size of the marker.
        /// </summary>
        ///
        /// <remarks>
        /// value between 2 and 72.
        /// </remarks>
        ///
        /// <value>
        /// The size of the marker.
        /// </value>
        [Obsolete("Please use Marker.Size")]
        public int MarkerSize
        {
            get
            {

                var size = Marker.Size;
                if (size == 0)
                {
                    return 5;
                }
                else
                {
                    return size;
                }
            }
            set
            {
                Marker.Size = value;
            }
        }
        /// <summary>
        /// Marker color.
        /// </summary>
        /// <value>
        /// The color of the Marker.
        /// </value>
        [Obsolete("Please use Marker.Fill")]
        public Color MarkerColor
        {
            get
            {
                if (Marker.Fill.Style == eFillStyle.SolidFill && Marker.Fill.SolidFill.Color.ColorType == eDrawingColorType.Rgb)
                {
                    return Marker.Fill.Color;
                }
                else
                {
                    return Color.Black;
                }
            }
            set
            {
                Marker.Fill.Color=value;
            }
        }

        /// <summary>
        /// Gets or sets the width of the line in pt.
        /// </summary>
        ///
        /// <value>
        /// The width of the line.
        /// </value>
        [Obsolete("Please use Border.Width")]
        public double LineWidth
        {
            get
            {
                var width = Border.Width;
                if (width == 0)
                {
                    return 2.25;
                }
                else
                {
                    return width;
                }
            }
            set
            {
                Border.Width = value;
            }
        }
        /// <summary>
        /// Marker Line color.
        /// (not to be confused with LineColor)
        /// </summary>
        ///
        /// <value>
        /// The color of the Marker line.
        /// </value>
        [Obsolete("Please use Marker.Border.Fill.Color")]
        public Color MarkerLineColor
        {
            get
            {                
                if (Marker.Border.Fill.Style==eFillStyle.SolidFill && Marker.Border.Fill.SolidFill.Color.ColorType==eDrawingColorType.Rgb)
                {
                    return Marker.Border.Fill.Color;
                }
                else
                {
                    return Color.Black;
                }
            }
            set
            {
                Marker.Border.Fill.Color = value;
            }
        }
    }
}
