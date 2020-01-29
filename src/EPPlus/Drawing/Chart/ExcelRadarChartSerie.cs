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
using OfficeOpenXml.Drawing.Interfaces;
using System;
using System.Globalization;
using System.Xml;
namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// A serie for a scatter chart
    /// </summary>
    public sealed class ExcelRadarChartSerie : ExcelChartSerie, IDrawingSerieDataLabel, IDrawingChartMarker, IDrawingChartDataPoints
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="chart">The chart</param>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="node">Topnode</param>
        /// <param name="isPivot">Is pivotchart</param>
        internal ExcelRadarChartSerie(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot) :
            base(chart, ns, node, isPivot)
        {
            if (chart.ChartType == eChartType.RadarMarkers)
            {
                Marker.Style = eMarkerStyle.Square;
            }
            else if(chart.ChartType == eChartType.Radar)
            {
                Marker.Style = eMarkerStyle.None;
            }
        }
        ExcelChartSerieDataLabel _DataLabel = null;
        /// <summary>
        /// Datalabel
        /// </summary>
        public ExcelChartSerieDataLabel DataLabel
        {
            get
            {
                if (_DataLabel == null)
                {
                    _DataLabel = new ExcelChartSerieDataLabel(_chart, NameSpaceManager, TopNode, SchemaNodeOrder);
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
                return TopNode.SelectSingleNode("c:dLbls", NameSpaceManager) != null;
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
                //if (IsMarkersAllowed() == false) return null;

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
                return ExistNode("c:marker");
            }
            return false;
        }
        private bool IsMarkersAllowed()
        {
            if (_chart.ChartType == eChartType.RadarMarkers)
            {
                return true;
            }
            return false;
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

        const string MARKERSIZE_PATH = "c:marker/c:size/@val";
        /// <summary>
        /// The size of a markers
        /// </summary>
        [Obsolete("Please use Marker.Size")]
        public int MarkerSize
        {
            get
            {
                return GetXmlNodeInt(MARKERSIZE_PATH);
            }
            set
            {
                if (value < 2 && value > 72)
                {
                    throw (new ArgumentOutOfRangeException("MarkerSize out of range. Range from 2-72 allowed."));
                }
                SetXmlNodeString(MARKERSIZE_PATH, value.ToString(CultureInfo.InvariantCulture));
            }
        }

    }
}
