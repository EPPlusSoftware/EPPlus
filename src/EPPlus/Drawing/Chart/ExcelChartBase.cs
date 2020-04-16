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
using System.IO;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing.Style.ThreeD;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Base class for Chart object.
    /// </summary>
    public abstract class ExcelChartBase : ExcelDrawing, IDrawingStyle, IStyleMandatoryProperties, IPictureRelationDocument
    {
        internal bool _isChartEx;
        internal const string topPath = "c:chartSpace";
        internal const string plotAreaPath = "c:chart/c:plotArea";
        //string _chartPath;
        internal ExcelChartAxis[] _axis;
        internal Dictionary<string, HashInfo> _hashes;
        /// <summary>
        /// The Xml helper for the chart xml
        /// </summary>
        protected internal XmlHelper _chartXmlHelper;
        internal ExcelChartBase _topChart = null;
        #region "Constructors"
        internal ExcelChartBase(ExcelDrawings drawings, XmlNode node, eChartType? type, bool isPivot, ExcelGroupShape parent, string drawingPath= "xdr:graphicFrame", string nvPrPath = "xdr:nvGraphicFramePr/xdr:cNvPr") :
            base(drawings, node, drawingPath, nvPrPath, parent)
        {            
        }
        internal ExcelChartBase(ExcelDrawings drawings, XmlNode drawingsNode, eChartType? type, ExcelChartBase topChart, ExcelPivotTable PivotTableSource, XmlDocument chartXml = null, ExcelGroupShape parent=null, string drawingPath = "xdr:graphicFrame", string nvPrPath = "xdr:nvGraphicFramePr/xdr:cNvPr") :
            base(drawings, drawingsNode, drawingPath, nvPrPath, parent)
        {            
        }
        internal ExcelChartBase(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent, string drawingPath = "xdr:graphicFrame", string nvPrPath = "xdr:nvGraphicFramePr/xdr:cNvPr") :
           base(drawings, node, drawingPath, nvPrPath, parent)
        {
        }
        internal ExcelChartBase(ExcelChart topChart, XmlNode chartNode, ExcelGroupShape parent, string drawingPath = "xdr:graphicFrame", string nvPrPath = "xdr:nvGraphicFramePr/xdr:cNvPr") :
            base(topChart._drawings, topChart.TopNode, drawingPath, nvPrPath, parent)
        {
        }
        #endregion
        internal ExcelChartStyleManager _styleManager = null;
        /// <summary>
        /// Manage style settings for the chart
        /// </summary>
        public ExcelChartStyleManager StyleManager
        {
            get
            {
                if (_styleManager == null)
                {
                    _styleManager = new ExcelChartStyleManager(NameSpaceManager, this);
                }
                return _styleManager;
            }
        }
        private bool HasPrimaryAxis()
        {
            if (_plotArea.ChartTypes.Count == 1)
            {
                return false;
            }
            foreach (var chart in _plotArea.ChartTypes)
            {
                if (chart != this)
                {
                    if (chart.UseSecondaryAxis == false && chart.IsTypePieDoughnut() == false)
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        internal abstract void AddAxis();
        bool _secondaryAxis = false;
        /// <summary>
        /// If true the charttype will use the secondary axis.
        /// The chart must contain a least one other charttype that uses the primary axis.
        /// </summary>
        public bool UseSecondaryAxis
        {
            get
            {
                return _secondaryAxis;
            }
            set
            {
                if (_secondaryAxis != value)
                {
                    if (value)
                    {
                        if (IsTypePieDoughnut())
                        {
                            throw (new Exception("Pie charts do not support axis"));
                        }
                        else if (HasPrimaryAxis() == false)
                        {
                            throw (new Exception("Can't set to secondary axis when no serie uses the primary axis"));
                        }
                        if (Axis.Length == 2)
                        {
                            AddAxis();
                        }
                        var nl = ChartNode.SelectNodes("c:axId", NameSpaceManager);
                        nl[0].Attributes["val"].Value = Axis[2].Id;
                        nl[1].Attributes["val"].Value = Axis[3].Id;
                        XAxis = Axis[2];
                        YAxis = Axis[3];
                    }
                    else
                    {
                        var nl = ChartNode.SelectNodes("c:axId", NameSpaceManager);
                        nl[0].Attributes["val"].Value = Axis[0].Id;
                        nl[1].Attributes["val"].Value = Axis[1].Id;
                        XAxis = Axis[0];
                        YAxis = Axis[1];
                    }
                    _secondaryAxis = value;
                }
            }
        }
        #region "Properties"
        /// <summary>
        /// Reference to the worksheet
        /// </summary>
        public ExcelWorksheet WorkSheet { get; internal set; }
        /// <summary>
        /// The chart xml document
        /// </summary>
        public XmlDocument ChartXml { get; internal set; }
        /// <summary>
        /// Type of chart
        /// </summary>
        public eChartType ChartType { get; internal set; }
        /// <summary>
        /// The chart element
        /// </summary>
        internal protected XmlNode _chartNode = null;
        internal XmlNode ChartNode
        {
            get
            {
                return _chartNode;
            }
        }
        /// <summary>
        /// Titel of the chart
        /// </summary>
        public ExcelChartTitle Title
        {
            get
            {
                if (_title == null)
                {
                    _title = new ExcelChartTitle(this, NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart", NameSpaceManager));
                }
                return _title;
            }
        }
        /// <summary>
        /// True if the chart has a title
        /// </summary>
        public bool HasTitle
        {
            get
            {
                return ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:title", NameSpaceManager) != null;
            }
        }
        /// <summary>
        /// If the chart has a legend
        /// </summary>
        public bool HasLegend
        {
            get
            {
                return ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:legend", NameSpaceManager) != null;
            }
        }
        /// <summary>
        /// Remove the title from the chart
        /// </summary>
        public void DeleteTitle()
        {
            _title = null;
            _chartXmlHelper.DeleteNode("../../c:title");
        }
        /// <summary>
        /// Chart series
        /// </summary>
        public virtual ExcelChartSeries<ExcelChartSerieBase> Series { get; } = new ExcelChartSeries<ExcelChartSerieBase>();
        /// <summary>
        /// An array containg all axis of all Charttypes
        /// </summary>
        public ExcelChartAxis[] Axis
        {
            get
            {
                return _axis;
            }
        }
        /// <summary>
        /// The X Axis
        /// </summary>
        public ExcelChartAxis XAxis
        {
            get;
            private set;
        }
        /// <summary>
        /// The Y Axis
        /// </summary>
        public ExcelChartAxis YAxis
        {
            get;
            private set;
        }
        /// <summary>
        /// The build-in chart styles. 
        /// </summary>
        public eChartStyle Style
        {
            get
            {
                XmlNode node = ChartXml.SelectSingleNode("c:chartSpace/c:style/@val", NameSpaceManager);
                if (node == null)
                {
                    return eChartStyle.None;
                }
                else
                {
                    int v;
                    if (int.TryParse(node.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out v))
                    {
                        return (eChartStyle)v;
                    }
                    else
                    {
                        return eChartStyle.None;
                    }
                }

            }
            set
            {
                if (value == eChartStyle.None)
                {
                    XmlElement element = ChartXml.SelectSingleNode("c:chartSpace/c:style", NameSpaceManager) as XmlElement;
                    if (element != null)
                    {
                        element.ParentNode.RemoveChild(element);
                    }
                }
                else
                {
                    if (!_chartXmlHelper.ExistNode("../../../c:style"))
                    {
                        XmlElement element = ChartXml.CreateElement("c:style", ExcelPackage.schemaChart);
                        element.SetAttribute("val", ((int)value).ToString());
                        XmlElement parent = ChartXml.SelectSingleNode("c:chartSpace", NameSpaceManager) as XmlElement;
                        parent.InsertBefore(element, parent.SelectSingleNode("c:chart", NameSpaceManager));
                    }
                    else
                    {
                        _chartXmlHelper.SetXmlNodeString("../../../ c:style/@val", ((int)value).ToString(CultureInfo.InvariantCulture));
                    }                    
                }
            }
        }
        ExcelChartPlotArea _plotArea = null;
        /// <summary>
        /// Plotarea
        /// </summary>
        public ExcelChartPlotArea PlotArea
        {
            get
            {
                if (_plotArea == null)
                {
                    _plotArea = new ExcelChartPlotArea(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:plotArea", NameSpaceManager), this);
                }
                return _plotArea;
            }
        }
        ExcelChartLegend _legend = null;
        /// <summary>
        /// Legend
        /// </summary>
        public ExcelChartLegend Legend
        {
            get
            {
                if (_legend == null)
                {
                    _legend = new ExcelChartLegend(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:legend", NameSpaceManager), this);
                }
                return _legend;
            }

        }
        ExcelDrawingBorder _border = null;
        /// <summary>
        /// Border
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(this, NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace", NameSpaceManager), "c:spPr/a:ln", _chartXmlHelper.SchemaNodeOrder);
                }
                return _border;
            }
        }
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// Access to Fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(this, NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace", NameSpaceManager), "c:spPr", _chartXmlHelper.SchemaNodeOrder);
                }
                return _fill;
            }
        }
        ExcelDrawingEffectStyle _effect = null;
        /// <summary>
        /// Effects
        /// </summary>
        public ExcelDrawingEffectStyle Effect
        {
            get
            {
                if (_effect == null)
                {
                    _effect = new ExcelDrawingEffectStyle(this, NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace", NameSpaceManager), "c:spPr/a:effectLst", _chartXmlHelper.SchemaNodeOrder);
                }
                return _effect;
            }
        }
        ExcelDrawing3D _threeD = null;
        /// <summary>
        /// 3D properties
        /// </summary>
        public ExcelDrawing3D ThreeD
        {
            get
            {
                if (_threeD == null)
                {
                    _threeD = new ExcelDrawing3D(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace", NameSpaceManager), "c:spPr", _chartXmlHelper.SchemaNodeOrder);
                }
                return _threeD;
            }
        }
        ExcelTextFont _font = null;
        /// <summary>
        /// Access to font properties
        /// </summary>
        public ExcelTextFont Font
        {
            get
            {
                if (_font == null)
                {
                    _font = new ExcelTextFont(this, NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace", NameSpaceManager), "c:txPr/a:p/a:pPr/a:defRPr", _chartXmlHelper.SchemaNodeOrder);
                }
                return _font;
            }
        }
        ExcelTextBody _textBody = null;
        /// <summary>
        /// Access to text body properties
        /// </summary>
        public ExcelTextBody TextBody
        {
            get
            {
                if (_textBody == null)
                {
                    _textBody = new ExcelTextBody(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace", NameSpaceManager), "c:txPr/a:bodyPr", _chartXmlHelper.SchemaNodeOrder);
                }
                return _textBody;
            }
        }

        /// <summary>
        /// If the chart is a pivochart this is the pivotable used as source.
        /// </summary>
        public ExcelPivotTable PivotTableSource
        {
            get;
            protected set;
        }
        void IDrawingStyleBase.CreatespPr()
        {
            _chartXmlHelper.CreatespPrNode("../../../c:spPr");
        }
        internal Packaging.ZipPackagePart Part { get; set; }
        /// <summary>
        /// Package internal URI
        /// </summary>
        internal Uri UriChart { get; set; }
        internal new string Id
        {
            get { return ""; }
        }
        ExcelChartTitle _title = null;
        #endregion
        #region "Chart type functions
        /// <summary>
        /// Returns true if the chart is a 3D chart
        /// </summary>
        /// <param name="chartType">The charttype to tests</param>
        /// <returns>True if the chart is a 3D chart</returns>
        internal static bool IsType3D(eChartType chartType)
        {
            return chartType == eChartType.Area3D ||
                            chartType == eChartType.AreaStacked3D ||
                            chartType == eChartType.AreaStacked1003D ||
                            chartType == eChartType.BarClustered3D ||
                            chartType == eChartType.BarStacked3D ||
                            chartType == eChartType.BarStacked1003D ||
                            chartType == eChartType.Column3D ||
                            chartType == eChartType.ColumnClustered3D ||
                            chartType == eChartType.ColumnStacked3D ||
                            chartType == eChartType.ColumnStacked1003D ||
                            chartType == eChartType.Line3D ||
                            chartType == eChartType.Pie3D ||
                            chartType == eChartType.PieExploded3D ||
                            chartType == eChartType.ConeBarClustered ||
                            chartType == eChartType.ConeBarStacked ||
                            chartType == eChartType.ConeBarStacked100 ||
                            chartType == eChartType.ConeCol ||
                            chartType == eChartType.ConeColClustered ||
                            chartType == eChartType.ConeColStacked ||
                            chartType == eChartType.ConeColStacked100 ||
                            chartType == eChartType.CylinderBarClustered ||
                            chartType == eChartType.CylinderBarStacked ||
                            chartType == eChartType.CylinderBarStacked100 ||
                            chartType == eChartType.CylinderCol ||
                            chartType == eChartType.CylinderColClustered ||
                            chartType == eChartType.CylinderColStacked ||
                            chartType == eChartType.CylinderColStacked100 ||
                            chartType == eChartType.PyramidBarClustered ||
                            chartType == eChartType.PyramidBarStacked ||
                            chartType == eChartType.PyramidBarStacked100 ||
                            chartType == eChartType.PyramidCol ||
                            chartType == eChartType.PyramidColClustered ||
                            chartType == eChartType.PyramidColStacked ||
                            chartType == eChartType.PyramidColStacked100 ||
                            chartType == eChartType.Surface ||
                            chartType == eChartType.SurfaceTopView ||
                            chartType == eChartType.SurfaceTopViewWireframe ||
                            chartType == eChartType.SurfaceWireframe;
        }

        /// <summary>
        /// Returns true if the chart is a 3D chart
        /// </summary>
        /// <returns>True if the chart is a 3D chart</returns>
        internal protected bool IsType3D()
        {
            return IsType3D(ChartType);
        }
        /// <summary>
        /// Returns true if the chart is a line chart
        /// </summary>
        /// <returns>True if the chart is a line chart</returns>
        protected internal bool IsTypeLine()
        {
            return ChartType == eChartType.Line ||
                    ChartType == eChartType.LineMarkers ||
                    ChartType == eChartType.LineMarkersStacked100 ||
                    ChartType == eChartType.LineStacked ||
                    ChartType == eChartType.LineStacked100 ||
                    ChartType == eChartType.Line3D;
        }
        /// <summary>
        /// Returns true if the chart is a radar chart
        /// </summary>
        /// <returns>True if the chart is a radar chart</returns>
        protected internal bool IsTypeRadar()
        {
            return ChartType == eChartType.Radar ||
                   ChartType == eChartType.RadarFilled ||
                   ChartType == eChartType.RadarMarkers;
        }

        /// <summary>
        /// Returns true if the chart is a scatter chart
        /// </summary>
        /// <returns>True if the chart is a scatter chart</returns>
        protected internal bool IsTypeScatter()
        {
            return ChartType == eChartType.XYScatter ||
                    ChartType == eChartType.XYScatterLines ||
                    ChartType == eChartType.XYScatterLinesNoMarkers ||
                    ChartType == eChartType.XYScatterSmooth ||
                    ChartType == eChartType.XYScatterSmoothNoMarkers;
        }
        /// <summary>
        /// Returns true if the chart is a bubble chart
        /// </summary>
        /// <returns>True if the chart is a bubble chart</returns>
        protected internal bool IsTypeBubble()
        {
            return ChartType == eChartType.Bubble ||
                    ChartType == eChartType.Bubble3DEffect;
        }
        /// <summary>
        /// Returns true if the chart is a sureface chart
        /// </summary>
        /// <returns>True if the chart is a sureface chart</returns>
        protected bool IsTypeSurface()
        {
            return ChartType == eChartType.Surface ||
                   ChartType == eChartType.SurfaceTopView ||
                   ChartType == eChartType.SurfaceTopViewWireframe ||
                   ChartType == eChartType.SurfaceWireframe;
        }
        /// <summary>
        /// Returns true if the chart is a sureface chart
        /// </summary>
        /// <returns>True if the chart is a sureface chart</returns>
        internal protected bool HasThirdAxis()
        {
            return IsTypeSurface() || ChartType == eChartType.Line3D;
        }
        /// <summary>
        /// Returns true if the chart has shapes, like bars and columns
        /// </summary>
        /// <returns>True if the chart has shapes</returns>
        protected internal bool IsTypeShape()
        {
            return ChartType == eChartType.BarClustered3D ||
                    ChartType == eChartType.BarStacked3D ||
                    ChartType == eChartType.BarStacked1003D ||
                    ChartType == eChartType.BarClustered3D ||
                    ChartType == eChartType.BarStacked3D ||
                    ChartType == eChartType.BarStacked1003D ||
                    ChartType == eChartType.Column3D ||
                    ChartType == eChartType.ColumnClustered3D ||
                    ChartType == eChartType.ColumnStacked3D ||
                    ChartType == eChartType.ColumnStacked1003D ||
                    //ChartType == eChartType.3DPie ||
                    //ChartType == eChartType.3DPieExploded ||
                    //ChartType == eChartType.Bubble3DEffect ||
                    ChartType == eChartType.ConeBarClustered ||
                    ChartType == eChartType.ConeBarStacked ||
                    ChartType == eChartType.ConeBarStacked100 ||
                    ChartType == eChartType.ConeCol ||
                    ChartType == eChartType.ConeColClustered ||
                    ChartType == eChartType.ConeColStacked ||
                    ChartType == eChartType.ConeColStacked100 ||
                    ChartType == eChartType.CylinderBarClustered ||
                    ChartType == eChartType.CylinderBarStacked ||
                    ChartType == eChartType.CylinderBarStacked100 ||
                    ChartType == eChartType.CylinderCol ||
                    ChartType == eChartType.CylinderColClustered ||
                    ChartType == eChartType.CylinderColStacked ||
                    ChartType == eChartType.CylinderColStacked100 ||
                    ChartType == eChartType.PyramidBarClustered ||
                    ChartType == eChartType.PyramidBarStacked ||
                    ChartType == eChartType.PyramidBarStacked100 ||
                    ChartType == eChartType.PyramidCol ||
                    ChartType == eChartType.PyramidColClustered ||
                    ChartType == eChartType.PyramidColStacked ||
                    ChartType == eChartType.PyramidColStacked100; //||
                                                                  //ChartType == eChartType.Doughnut ||
                                                                  //ChartType == eChartType.DoughnutExploded;
        }
        /// <summary>
        /// Returns true if the chart is of type stacked percentage
        /// </summary>
        /// <returns>True if the chart is of type stacked percentage</returns>
        protected internal bool IsTypePercentStacked()
        {
            return ChartType == eChartType.AreaStacked100 ||
                           ChartType == eChartType.BarStacked100 ||
                           ChartType == eChartType.BarStacked1003D ||
                           ChartType == eChartType.ColumnStacked100 ||
                           ChartType == eChartType.ColumnStacked1003D ||
                           ChartType == eChartType.ConeBarStacked100 ||
                           ChartType == eChartType.ConeColStacked100 ||
                           ChartType == eChartType.CylinderBarStacked100 ||
                           ChartType == eChartType.CylinderColStacked ||
                           ChartType == eChartType.LineMarkersStacked100 ||
                           ChartType == eChartType.LineStacked100 ||
                           ChartType == eChartType.PyramidBarStacked100 ||
                           ChartType == eChartType.PyramidColStacked100;
        }
        /// <summary>
        /// Returns true if the chart is of type stacked 
        /// </summary>
        /// <returns>True if the chart is of type stacked</returns>
        protected internal bool IsTypeStacked()
        {
            return ChartType == eChartType.AreaStacked ||
                           ChartType == eChartType.AreaStacked3D ||
                           ChartType == eChartType.BarStacked ||
                           ChartType == eChartType.BarStacked3D ||
                           ChartType == eChartType.ColumnStacked3D ||
                           ChartType == eChartType.ColumnStacked ||
                           ChartType == eChartType.ConeBarStacked ||
                           ChartType == eChartType.ConeColStacked ||
                           ChartType == eChartType.CylinderBarStacked ||
                           ChartType == eChartType.CylinderColStacked ||
                           ChartType == eChartType.LineMarkersStacked ||
                           ChartType == eChartType.LineStacked ||
                           ChartType == eChartType.PyramidBarStacked ||
                           ChartType == eChartType.PyramidColStacked;
        }
        /// <summary>
        /// Returns true if the chart is of type clustered
        /// </summary>
        /// <returns>True if the chart is of type clustered</returns>
        protected bool IsTypeClustered()
        {
            return ChartType == eChartType.BarClustered ||
                           ChartType == eChartType.BarClustered3D ||
                           ChartType == eChartType.ColumnClustered3D ||
                           ChartType == eChartType.ColumnClustered ||
                           ChartType == eChartType.ConeBarClustered ||
                           ChartType == eChartType.ConeColClustered ||
                           ChartType == eChartType.CylinderBarClustered ||
                           ChartType == eChartType.CylinderColClustered ||
                           ChartType == eChartType.PyramidBarClustered ||
                           ChartType == eChartType.PyramidColClustered;
        }
        /// <summary>
        /// Returns true if the chart is a pie or Doughnut chart
        /// </summary>
        /// <returns>True if the chart is a pie or Doughnut chart</returns>
        protected internal bool IsTypePieDoughnut()
        {
            return IsTypePie() || IsTypeDoughnut();
        }
        /// <summary>
        /// Returns true if the chart is a Doughnut chart
        /// </summary>
        /// <returns>True if the chart is a Doughnut chart</returns>
        protected internal bool IsTypeDoughnut()
        {
            return ChartType == eChartType.Doughnut ||
                           ChartType == eChartType.DoughnutExploded;
        }
        /// <summary>
        /// Returns true if the chart is a pie chart
        /// </summary>
        /// <returns>True if the chart is a pie chart</returns>
        protected internal bool IsTypePie()
        {
            return ChartType == eChartType.Pie ||
                           ChartType == eChartType.PieExploded ||
                           ChartType == eChartType.PieOfPie ||
                           ChartType == eChartType.Pie3D ||
                           ChartType == eChartType.PieExploded3D ||
                           ChartType == eChartType.BarOfPie;
        }
        #endregion
        internal void InitChartTheme(int fallBackStyleId)
        {
            var styleId = fallBackStyleId + 100;
            XmlElement el = (XmlElement)_chartXmlHelper.CreateNode("../../../mc:AlternateContent/mc:Choice");
            el.SetAttribute("xmlns:c14", ExcelPackage.schemaChart14);
            _chartXmlHelper.SetXmlNodeString("../../../mc:AlternateContent/mc:Choice/@Requires", "c14");
            _chartXmlHelper.SetXmlNodeString("../../../mc:AlternateContent/mc:Choice/c14:style/@val", styleId.ToString(CultureInfo.InvariantCulture));
            _chartXmlHelper.SetXmlNodeString("../../../mc:AlternateContent/mc:Fallback/c:style/@val", fallBackStyleId.ToString(CultureInfo.InvariantCulture));
        }
        public abstract bool VaryColors { get; set; }
        /// <summary>
        /// Formatting for the floor of a 3D chart. 
        /// <note type="note">This property is null for non 3D charts</note>
        /// </summary>
        public ExcelChartSurface Floor { get; protected set; } = null;
        /// <summary>
        /// Formatting for the sidewall of a 3D chart. 
        /// <note type="note">This property is null for non 3D charts</note>
        /// </summary>
        public ExcelChartSurface SideWall { get; protected set; } = null;
        /// <summary>
        /// Formatting for the backwall of a 3D chart. 
        /// <note type="note">This property is null for non 3D charts</note>
        /// </summary>
        public ExcelChartSurface BackWall { get; protected set; } = null; 
        internal override void DeleteMe()
        {
            try
            {
                Part.Package.DeletePart(UriChart);
            }
            catch (Exception ex)
            {
                throw (new InvalidDataException("EPPlus internal error when deleting chart.", ex));
            }
            base.DeleteMe();
        }
        void IStyleMandatoryProperties.SetMandatoryProperties()
        {
            _chartXmlHelper.CreatespPrNode("../c:spPr");
        }
        ExcelPackage IPictureRelationDocument.Package => _drawings._package;

        Dictionary<string, HashInfo> IPictureRelationDocument.Hashes => _hashes;

        ZipPackagePart IPictureRelationDocument.RelatedPart => Part;

        Uri IPictureRelationDocument.RelatedUri => UriChart;
    }
}
