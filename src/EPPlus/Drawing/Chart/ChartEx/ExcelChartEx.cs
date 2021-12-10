/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/15/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.Constants;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// Base class for all extention charts
    /// </summary>
    public abstract class ExcelChartEx : ExcelChart
    {
        internal ExcelChartEx(ExcelDrawings drawings, XmlNode node, ExcelGroupShape parent) : 
            base(drawings, node, parent, "mc:AlternateContent/mc:Choice/xdr:graphicFrame")
        {
            ChartType = GetChartType(node, drawings.NameSpaceManager);
            Init();
        }

        internal ExcelChartEx(ExcelDrawings drawings, XmlNode drawingsNode, eChartType? type, XmlDocument chartXml = null, ExcelGroupShape parent = null) :
            base(drawings, drawingsNode, chartXml, parent, "mc:AlternateContent/mc:Choice/xdr:graphicFrame")
       {
            ChartType = type.Value;
            CreateNewChart(drawings, chartXml, type);
            Init();
        }
        internal ExcelChartEx(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent=null) :
            base(drawings, node, chartXml, parent, "mc:AlternateContent/mc:Choice/xdr:graphicFrame")
        {
            UriChart = uriChart;
            Part = part;
            _chartNode = chartNode;
            _chartXmlHelper = XmlHelperFactory.Create(drawings.NameSpaceManager, chartNode);
            ChartType = GetChartType(chartNode, drawings.NameSpaceManager);
            Init();
        }
        internal void LoadAxis()
        {
            var l = new List<ExcelChartAxis>();            
            foreach (XmlNode axNode in _chartXmlHelper.GetNodes("cx:plotArea/cx:axis"))
            {
                l.Add(new ExcelChartExAxis(this, NameSpaceManager, axNode));
            }
            _axis = l.ToArray();
            _exAxis = null;
            if(Axis.Length>0)
            {
                if(Axis[1].AxisType==eAxisType.Cat)
                {
                    XAxis = Axis[1];
                    YAxis = Axis[0];
                }
                else
                {
                    XAxis = Axis[0];
                    YAxis = Axis[1];
                }
            }
        }
        private void Init()
        {
            _isChartEx = true;
            _chartXmlHelper.SchemaNodeOrder = new string[] { "chartData", "chart", "spPr", "txPr", "clrMapOvr", "fmtOvrs", "title", "plotArea","plotAreaRegion","axis", "legend", "printSettings" };
            base.Series.Init(this, NameSpaceManager, _chartNode, false);
            Series.Init(this, NameSpaceManager, _chartNode, false, base.Series._list);
            LoadAxis();
        }

        private void CreateNewChart(ExcelDrawings drawings, XmlDocument chartXml = null, eChartType? type = null)
        {
            XmlElement graphFrame = TopNode.OwnerDocument.CreateElement("mc","AlternateContent", ExcelPackage.schemaMarkupCompatibility);
            graphFrame.SetAttribute("xmlns:mc", ExcelPackage.schemaMarkupCompatibility);
            TopNode.AppendChild(graphFrame);
            graphFrame.InnerXml = string.Format("<mc:Choice xmlns:cx1=\"{1}\" Requires=\"cx1\"><xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"{0}\" name=\"\"><a:extLst><a:ext uri=\"{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}\"><a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{{9FE3C5B3-14FE-44E2-AB27-50960A44C7C4}}\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/drawing/2014/chartex\"><cx:chart xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"{0}\" name=\"\"/><xdr:cNvSpPr><a:spLocks noTextEdit=\"1\"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"3609974\" y=\"938212\"/><a:ext cx=\"5762625\" cy=\"2743200\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val=\"white\"/></a:solidFill><a:ln w=\"1\"><a:solidFill><a:prstClr val=\"green\"/></a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\"/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\" sz=\"1100\"/><a:t>This chart isn't available in your version of Excel. Editing this shape or saving this workbook into a different file format will permanently break the chart.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>", _id,GetChartExNameSpace(type??eChartType.Sunburst));
            TopNode.AppendChild(TopNode.OwnerDocument.CreateElement("clientData", ExcelPackage.schemaSheetDrawings));

            var package = drawings.Worksheet._package.ZipPackage;
            UriChart = GetNewUri(package, "/xl/charts/chartex{0}.xml");

            if (chartXml == null)
            {
                ChartXml = new XmlDocument
                {
                    PreserveWhitespace = ExcelPackage.preserveWhitespace
                };
                LoadXmlSafe(ChartXml, ChartStartXml(type.Value), Encoding.UTF8);
            }
            else
            {
                ChartXml = chartXml;
            }

            // save it to the package
            Part = package.CreatePart(UriChart, ContentTypes.contentTypeChartEx, _drawings._package.Compression);

            StreamWriter streamChart = new StreamWriter(Part.GetStream(FileMode.Create, FileAccess.Write));
            ChartXml.Save(streamChart);
            streamChart.Close();
            package.Flush();

            var chartRelation = drawings.Part.CreateRelationship(UriHelper.GetRelativeUri(drawings.UriDrawing, UriChart), Packaging.TargetMode.Internal, ExcelPackage.schemaChartExRelationships);
            graphFrame.SelectSingleNode("mc:Choice/xdr:graphicFrame/a:graphic/a:graphicData/cx:chart", NameSpaceManager).Attributes["r:id"].Value = chartRelation.Id;
            package.Flush();
            _chartNode = ChartXml.SelectSingleNode("cx:chartSpace/cx:chart", NameSpaceManager);
            _chartXmlHelper = XmlHelperFactory.Create(NameSpaceManager, _chartNode);
            GetPositionSize();
        }

        private string GetChartExNameSpace(eChartType type)
        {
            switch(type)
            {
                case eChartType.RegionMap:
                    return "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex";
                case eChartType.Funnel:
                    return "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex";
                default:
                    return "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex";
            }
            
        }

        private string ChartStartXml(eChartType type)
        {
            StringBuilder xml = new StringBuilder();

            xml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            xml.Append("<cx:chartSpace xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" >");
            xml.Append("<cx:chart><cx:title overlay=\"0\" align=\"ctr\" pos=\"t\"/><cx:plotArea><cx:plotAreaRegion></cx:plotAreaRegion></cx:plotArea></cx:chart>");
            xml.Append("</cx:chartSpace>");

            return xml.ToString();
        }

        private static void AddData(StringBuilder xml)
        {
            xml.Append("<cx:chartData><cx:data id=\"0\"><cx:strDim type=\"cat\"><cx:f dir=\"row\">_xlchart.v1.31</cx:f></cx:strDim><cx:numDim type=\"size\"><cx:f dir=\"row\">_xlchart.v1.32</cx:f></cx:numDim></cx:data><cx:data id=\"1\"><cx:strDim type=\"cat\"><cx:f dir=\"row\">_xlchart.v1.31</cx:f></cx:strDim><cx:numDim type=\"size\"><cx:f dir=\"row\">_xlchart.v1.33</cx:f></cx:numDim></cx:data></cx:chartData>");
        }

        internal override void AddAxis()
        {
            var l = new List<ExcelChartAxis>();
            foreach (XmlNode axNode in _chartXmlHelper.GetNodes("cx:plotArea/cx:axis"))
            {
                l.Add(new ExcelChartExAxis(this, NameSpaceManager, axNode));
            }
        }

        private static eChartType GetChartType(XmlNode node, XmlNamespaceManager nsm)
        {
            var layoutId = node.SelectSingleNode("cx:plotArea/cx:plotAreaRegion/cx:series[1]/@layoutId", nsm);
            if (layoutId == null) return eChartType.Treemap;
            switch (layoutId.Value)
            {
                case "clusteredColumn":
                    layoutId = node.SelectSingleNode("cx:plotArea/cx:plotAreaRegion/cx:series[@layoutId='paretoLine']", nsm);
                    if(layoutId==null)
                    {
                        return eChartType.Histogram;
                    }
                    else
                    {
                        return eChartType.Pareto;
                    }
                case "paretoLine":
                    return eChartType.Pareto;
                case "boxWhisker":
                    return eChartType.BoxWhisker;
                case "funnel":
                    return eChartType.Funnel;
                case "regionMap":
                    return eChartType.RegionMap;
                case "sunburst":
                    return eChartType.Sunburst;
                case "treemap":
                    return eChartType.Treemap;
                case "waterfall":
                    return eChartType.Waterfall;
                default:
                    throw new InvalidOperationException($"Unsupported layoutId in ChartEx Xml: {layoutId}");
            }
        }

        /// <summary>
        /// Delete the charts title
        /// </summary>
        public override void DeleteTitle()
        {
            _chartXmlHelper.DeleteNode("cx:title");
        }

        /// <summary>
        /// Plotarea properties
        /// </summary>
        public override ExcelChartPlotArea PlotArea
        {
            get
            {
                if (_plotArea==null)
                {
                    var node = _chartXmlHelper.GetNode("cx:plotArea");
                    _plotArea = new ExcelChartExPlotarea(NameSpaceManager, node, this);
                }
                return _plotArea;
            }
        }
        internal ExcelChartExAxis[] _exAxis = null;
        /// <summary>
        /// An array containg all axis of all Charttypes
        /// </summary>
        public new ExcelChartExAxis[] Axis
        {
            get
            {
                if(_exAxis==null)
                {
                    _exAxis=_axis.Select(x => (ExcelChartExAxis)x).ToArray();
                }
                return _exAxis;
            }
        }

        /// <summary>
        /// The titel of the chart
        /// </summary>
        public new ExcelChartExTitle Title
        {
            get
            {
                if (_title == null)
                {
                    return (ExcelChartExTitle)base.Title;
                }
                return (ExcelChartExTitle)_title;
            }
        }
        /// <summary>
        /// Legend
        /// </summary>
        public new ExcelChartExLegend Legend
        {
            get
            {
                if (_legend == null)
                {
                    return (ExcelChartExLegend)base.Legend;
                }
                return (ExcelChartExLegend)_legend;
            }
        }

        ExcelDrawingBorder _border = null;
        /// <summary>
        /// Border
        /// </summary>
        public override ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(this, NameSpaceManager, ChartXml.SelectSingleNode("cx:chartSpace", NameSpaceManager), "cx:spPr/a:ln", _chartXmlHelper.SchemaNodeOrder);
                }
                return _border;
            }
        }
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// Access to Fill properties
        /// </summary>
        public override ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(this, NameSpaceManager, ChartXml.SelectSingleNode("cx:chartSpace", NameSpaceManager), "cx:spPr", _chartXmlHelper.SchemaNodeOrder);
                }
                return _fill;
            }
        }
        ExcelDrawingEffectStyle _effect = null;
        /// <summary>
        /// Effects
        /// </summary>
        public override ExcelDrawingEffectStyle Effect
        {
            get
            {
                if (_effect == null)
                {
                    _effect = new ExcelDrawingEffectStyle(this, NameSpaceManager, ChartXml.SelectSingleNode("cx:chartSpace", NameSpaceManager), "cx:spPr/a:effectLst", _chartXmlHelper.SchemaNodeOrder);
                }
                return _effect;
            }
        }
        ExcelDrawing3D _threeD = null;
        /// <summary>
        /// 3D properties
        /// </summary>
        public override ExcelDrawing3D ThreeD
        {
            get
            {
                if (_threeD == null)
                {
                    _threeD = new ExcelDrawing3D(NameSpaceManager, ChartXml.SelectSingleNode("cx:chartSpace", NameSpaceManager), "cx:spPr", _chartXmlHelper.SchemaNodeOrder);
                }
                return _threeD;
            }
        }
        ExcelTextFont _font = null;
        /// <summary>
        /// Access to font properties
        /// </summary>
        public override ExcelTextFont Font
        {
            get
            {
                if (_font == null)
                {
                    _font = new ExcelTextFont(this, NameSpaceManager, ChartXml.SelectSingleNode("cx:chartSpace", NameSpaceManager), "cx:txPr/a:p/a:pPr/a:defRPr", _chartXmlHelper.SchemaNodeOrder);
                }
                return _font;
            }
        }
        ExcelTextBody _textBody = null;
        /// <summary>
        /// Access to text body properties
        /// </summary>
        public override ExcelTextBody TextBody
        {
            get
            {
                if (_textBody == null)
                {
                    _textBody = new ExcelTextBody(NameSpaceManager, ChartXml.SelectSingleNode("cx:chartSpace", NameSpaceManager), "cx:txPr/a:bodyPr", _chartXmlHelper.SchemaNodeOrder);
                }
                return _textBody;
            }
        }
        /// <summary>
        /// Chart series
        /// </summary>
        public new ExcelChartSeries<ExcelChartExSerie> Series { get; } = new ExcelChartSeries<ExcelChartExSerie>();
        /// <summary>
        /// Do not apply to Extension charts
        /// </summary>
        public override bool VaryColors
        {
            get 
            { 
                return false; 
            }
            set
            {
                throw new InvalidOperationException("VaryColors do not apply to Extended charts");
            }
        }
        /// <summary>
        /// Cannot be set for extension charts. Please use <see cref="ExcelChart.StyleManager"/>
        /// </summary>
        public override eChartStyle Style 
        {
            get;
            set;
        }
        /// <summary>
        /// If the chart has a title or not
        /// </summary>
        public override bool HasTitle
        {
            get
            {
                return _chartXmlHelper.ExistsNode("cx:title");
            }
        }

        /// <summary>
        /// If the chart has legend or not
        /// </summary>
        public override bool HasLegend
        {
            get
            {
                return _chartXmlHelper.ExistsNode("cx:legend");
            }
        }
        public override ExcelView3D View3D
        {
            get
            {
                return null;
            }
        }
        /// <summary>
        /// This property does not apply to extended charts.
        /// This property will always return eDisplayBlanksAs.Zero.
        /// Setting this property on an extended chart will result in an InvalidOperationException
        /// </summary>
        public override eDisplayBlanksAs DisplayBlanksAs 
        {
            get
            {
                return eDisplayBlanksAs.Zero;
            }
            set
            {
                throw new InvalidOperationException("DisplayBlanksAs do not apply to Extended charts");
            }
        }
        /// <summary>
        /// This property does not apply to extended charts.
        /// Setting this property on an extended chart will result in an InvalidOperationException
        /// </summary>
        public override bool RoundedCorners 
        {
            get
            {
                
                return false;
            }
            set
            {
                throw new InvalidOperationException("RoundedCorners do not apply to Extended charts");
            }
        }
        /// <summary>
        /// This property does not apply to extended charts.
        /// Setting this property on an extended chart will result in an InvalidOperationException
        /// </summary>
        public override bool ShowDataLabelsOverMaximum 
        {
            get
            {
                return false;
            }
            set
            {
                throw new InvalidOperationException("ShowHiddenData do not apply to Extended charts");
            }
        }
        /// <summary>
        /// This property does not apply to extended charts.
        /// Setting this property on an extended chart will result in an InvalidOperationException
        /// </summary>
        public override bool ShowHiddenData 
        {
            get
            {
                return false;
            }
            set
            {
                throw new InvalidOperationException("ShowHiddenData do not apply to Extended charts");
            }
        }
    }
}
