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
using OfficeOpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.Style
{
    /// <summary>
    /// Represents a style for a chart
    /// </summary>
    public class ExcelChartStyle : XmlHelper, IPictureRelationDocument
    {
        ExcelChartStyleManager _manager;
        Dictionary<string, HashInfo> _hashes = new Dictionary<string, HashInfo>();
        internal ExcelChartStyle(XmlNamespaceManager nsm, XmlNode topNode, ExcelChartStyleManager manager) : base(nsm, topNode)
        {
            _manager = manager;            
        }
        ExcelChartStyleEntry _axisTitle = null;
        /// <summary>
        /// Default formatting for an axis title.
        /// </summary>
        public ExcelChartStyleEntry AxisTitle
        {
            get
            {
                if (_axisTitle == null)
                {
                    _axisTitle = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:axisTitle", this);
                }
                return _axisTitle;
            }
        }
        ExcelChartStyleEntry _categoryAxis = null;
        /// <summary>
        /// Default formatting for a category axis
        /// </summary>
        public ExcelChartStyleEntry CategoryAxis
        {
            get
            {
                if (_categoryAxis == null)
                {
                    _categoryAxis = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:categoryAxis", this);
                }
                return _categoryAxis;
            }
        }
        ExcelChartStyleEntry _chartArea = null;
        /// <summary>
        /// Default formatting for a chart area
        /// </summary>
        public ExcelChartStyleEntry ChartArea
        {
            get
            {
                if (_chartArea == null)
                {
                    _chartArea = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:chartArea", this);
                }
                return _chartArea;
            }
        }
        ExcelChartStyleEntry _dataLabel = null;
        /// <summary>
        /// Default formatting for a data label
        /// </summary>
        public ExcelChartStyleEntry DataLabel
        {
            get
            {
                if (_dataLabel == null)
                {
                    _dataLabel = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:dataLabel", this);
                }
                return _dataLabel;
            }
        }
        ExcelChartStyleEntry _dataLabelCallout = null;
        /// <summary>
        /// Default formatting for a data label callout
        /// </summary>
        public ExcelChartStyleEntry DataLabelCallout
        {
            get
            {
                if (_dataLabelCallout == null)
                {
                    _dataLabelCallout = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:dataLabelCallout", this);
                }
                return _dataLabelCallout;
            }
        }
        ExcelChartStyleEntry _dataPoint = null;
        /// <summary>
        /// Default formatting for a data point on a 2-D chart of type column, bar,	filled radar, stock, bubble, pie, doughnut, area and 3-D bubble.
        /// </summary>
        public ExcelChartStyleEntry DataPoint
        {
            get
            {
                if (_dataPoint == null)
                {
                    _dataPoint = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:dataPoint", this);
                }
                return _dataPoint;
            }
        }
        ExcelChartStyleEntry _dataPoint3D = null;
        /// <summary>
        /// Default formatting for a data point on a 3-D chart of type column, bar, line, pie, area and surface.
        /// </summary>
        public ExcelChartStyleEntry DataPoint3D
        {
            get
            {
                if (_dataPoint3D == null)
                {
                    _dataPoint3D = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:dataPoint3D", this);
                }
                return _dataPoint3D;
            }
        }
        ExcelChartStyleEntry _dataPointLine = null;
        /// <summary>
        /// Default formatting for a data point on a 2-D chart of type line, scatter and radar
        /// </summary>
        public ExcelChartStyleEntry DataPointLine
        {
            get
            {
                if (_dataPointLine == null)
                {
                    _dataPointLine = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:dataPointLine", this);
                }
                return _dataPointLine;
            }
        }
        ExcelChartStyleEntry _dataPointMarker = null;
        /// <summary>
        /// Default formatting for a datapoint marker
        /// </summary>
        public ExcelChartStyleEntry DataPointMarker
        {
            get
            {
                if (_dataPointMarker == null)
                {
                    _dataPointMarker = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:dataPointMarker", this);
                }
                return _dataPointMarker;
            }
        }
        ExcelChartStyleMarkerLayout _dataPointMarkerLayout = null;
        /// <summary>
        /// Extended marker properties for a datapoint 
        /// </summary>
        public ExcelChartStyleMarkerLayout DataPointMarkerLayout
        {
            get
            {
                if (_dataPointMarkerLayout == null)
                {
                    var node = GetNode("cs:dataPointMarkerLayout");
                    if(node == null)
                    {
                        throw new InvalidOperationException("Invalid Chartstyle xml: dataPointMarkerLayout element missing");
                    }
                    _dataPointMarkerLayout = new ExcelChartStyleMarkerLayout(NameSpaceManager, node);
                }
                return _dataPointMarkerLayout;
            }
        }
        ExcelChartStyleEntry _dataPointWireframe = null;
        /// <summary>
        /// Default formatting for a datapoint on a surface wireframe chart
        /// </summary>
        public ExcelChartStyleEntry DataPointWireframe
        {
            get
            {
                if (_dataPointWireframe == null)
                {
                    _dataPointWireframe = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:dataPointWireframe", this);
                }
                return _dataPointWireframe;
            }
        }
        ExcelChartStyleEntry _dataTable = null;
        /// <summary>
        /// Default formatting for a Data table
        /// </summary>
        public ExcelChartStyleEntry DataTable
        {
            get
            {
                if (_dataTable == null)
                {
                    _dataTable = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:dataTable", this);
                }
                return _dataTable;
            }
        }
        ExcelChartStyleEntry _downBar = null;
        /// <summary>
        /// Default formatting for a downbar
        /// </summary>
        public ExcelChartStyleEntry DownBar
        {
            get
            {
                if (_downBar == null)
                {
                    _downBar = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:downBar", this);
                }
                return _downBar;
            }
        }
        ExcelChartStyleEntry _dropLine = null;
        /// <summary>
        /// Default formatting for a dropline
        /// </summary>
        public ExcelChartStyleEntry DropLine
        {
            get
            {
                if (_dropLine == null)
                {
                    _dropLine = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:dropLine", this);
                }
                return _dropLine;
            }
        }
        ExcelChartStyleEntry _errorBar = null;
        /// <summary>
        /// Default formatting for an errorbar
        /// </summary>
        public ExcelChartStyleEntry ErrorBar
        {
            get
            {
                if (_errorBar == null)
                {
                    _errorBar = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:errorBar", this);
                }
                return _errorBar;
            }
        }
        ExcelChartStyleEntry _floor = null;
        /// <summary>
        /// Default formatting for a floor
        /// </summary>
        public ExcelChartStyleEntry Floor
        {
            get
            {
                if (_floor == null)
                {
                    _floor = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:floor", this);
                }
                return _floor;
            }
        }
        ExcelChartStyleEntry _gridlineMajor = null;
        /// <summary>
        /// Default formatting for a major gridline
        /// </summary>
        public ExcelChartStyleEntry GridlineMajor
        {
            get
            {
                if (_gridlineMajor == null)
                {
                    _gridlineMajor = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:gridlineMajor", this);
                }
                return _gridlineMajor;
            }
        }
        ExcelChartStyleEntry _gridlineMinor = null;
        /// <summary>
        /// Default formatting for a minor gridline
        /// </summary>
        public ExcelChartStyleEntry GridlineMinor
        {
            get
            {
                if (_gridlineMinor == null)
                {
                    _gridlineMinor = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:gridlineMinor", this);
                }
                return _gridlineMinor;
            }
        }
        ExcelChartStyleEntry _hiLoLine = null;
        /// <summary>
        /// Default formatting for a high low line
        /// </summary>
        public ExcelChartStyleEntry HighLowLine
        {
            get
            {
                if (_hiLoLine == null)
                {
                    _hiLoLine = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:hiLoLine", this);
                }
                return _hiLoLine;
            }
        }
        ExcelChartStyleEntry _leaderLine = null;
        /// <summary>
        /// Default formatting for a leader line
        /// </summary>
        public ExcelChartStyleEntry LeaderLine
        {
            get
            {
                if (_leaderLine == null)
                {
                    _leaderLine = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:leaderLine", this);
                }
                return _leaderLine;
            }
        }
        /// <summary>
        /// Default formatting for a legend
        /// </summary>
        ExcelChartStyleEntry _legend = null;
        /// <summary>
        /// Default formatting for a chart legend
        /// </summary>
        public ExcelChartStyleEntry Legend
        {
            get
            {
                if (_legend == null)
                {
                    _legend = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:legend", this);
                }
                return _legend;
            }
        }
        ExcelChartStyleEntry _plotArea = null;
        /// <summary>
        /// Default formatting for a plot area in a 2D chart
        /// </summary>
        public ExcelChartStyleEntry PlotArea
        {
            get
            {
                if (_plotArea == null)
                {
                    _plotArea = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:plotArea", this);
                }
                return _plotArea;
            }
        }
        ExcelChartStyleEntry _plotArea3D = null;
        /// <summary>
        /// Default formatting for a plot area in a 3D chart
        /// </summary>
        public ExcelChartStyleEntry PlotArea3D
        {
            get
            {
                if (_plotArea3D == null)
                {
                    _plotArea3D = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:plotArea3D", this);
                }
                return _plotArea3D;
            }
        }
        ExcelChartStyleEntry _seriesAxis = null;
        /// <summary>
        /// Default formatting for a series axis
        /// </summary>
        public ExcelChartStyleEntry SeriesAxis
        {
            get
            {
                if (_seriesAxis == null)
                {
                    _seriesAxis = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:seriesAxis", this);
                }
                return _seriesAxis;
            }
        }
        ExcelChartStyleEntry _seriesLine = null;
        /// <summary>
        /// Default formatting for a series line
        /// </summary>
        public ExcelChartStyleEntry SeriesLine
        {
            get
            {
                if (_seriesLine == null)
                {
                    _seriesLine = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:seriesLine", this);
                }
                return _seriesLine;
            }
        }
        ExcelChartStyleEntry _title = null;
        /// <summary>
        /// Default formatting for a chart title
        /// </summary>
        public ExcelChartStyleEntry Title
        {
            get
            {
                if (_title == null)
                {
                    _title = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:title", this);
                }
                return _title;
            }
        }
        ExcelChartStyleEntry _trendline = null;
        /// <summary>
        /// Default formatting for a trend line
        /// </summary>
        public ExcelChartStyleEntry Trendline
        {
            get
            {
                if (_trendline == null)
                {
                    _trendline = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:trendline", this);
                }
                return _trendline;
            }
        }
        ExcelChartStyleEntry _trendlineLabel = null;
        /// <summary>
        /// Default formatting for a trend line label
        /// </summary>
        public ExcelChartStyleEntry TrendlineLabel
        {
            get
            {
                if (_trendlineLabel == null)
                {
                    _trendlineLabel = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:trendlineLabel", this);
                }
                return _trendlineLabel;
            }
        }
        ExcelChartStyleEntry _upBar = null;
        /// <summary>
        /// Default formatting for a up bar
        /// </summary>
        public ExcelChartStyleEntry UpBar
        {
            get
            {
                if (_upBar == null)
                {
                    _upBar = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:upBar", this);
                }
                return _upBar;
            }
        }
        ExcelChartStyleEntry _valueAxis = null;
        /// <summary>
        /// Default formatting for a value axis
        /// </summary>
        public ExcelChartStyleEntry ValueAxis
        {
            get
            {
                if (_valueAxis == null)
                {
                    _valueAxis = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:valueAxis", this);
                }
                return _valueAxis;
            }
        }
        ExcelChartStyleEntry _wall = null;
        /// <summary>
        /// Default formatting for a wall
        /// </summary>
        public ExcelChartStyleEntry Wall
        {
            get
            {
                if (_wall == null)
                {
                    _wall = new ExcelChartStyleEntry(NameSpaceManager, TopNode, "cs:wall", this);
                }
                return _wall;
            }
        }

        /// <summary>
        /// The id of the chart style
        /// </summary>
        public int Id
        {
            get
            {
                return GetXmlNodeInt("@id");
            }
            internal set
            {
                SetXmlNodeString("@id", value.ToString(CultureInfo.InvariantCulture));
            }
        }

        ExcelPackage IPictureRelationDocument.Package => _manager._chart._drawings._package;

        Dictionary<string, HashInfo> IPictureRelationDocument.Hashes => _hashes;

        ZipPackagePart IPictureRelationDocument.RelatedPart => _manager.StylePart;

        Uri IPictureRelationDocument.RelatedUri => _manager.StyleUri;
    }
}
