/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/15/2020         EPPlus Software AB       EPPlus 5.2
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
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Constants;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Base class for Chart object.
    /// </summary>
    public class ExcelChartStandard : ExcelChart
    {
        #region "Constructors"
        internal ExcelChartStandard(ExcelDrawings drawings, XmlNode node, eChartType? type, bool isPivot, ExcelGroupShape parent, string drawingPath = "xdr:graphicFrame", string nvPrPath = "xdr:nvGraphicFramePr/xdr:cNvPr") :
            base(drawings, node, parent, drawingPath, nvPrPath)
        {
            if (type.HasValue) ChartType = type.Value;
            CreateNewChart(drawings, null, null, type);

            Init(drawings, _chartNode);
            InitSeries(this, drawings.NameSpaceManager, _chartNode, isPivot);
            SetTypeProperties();
            LoadAxis();
        }
        internal ExcelChartStandard(ExcelDrawings drawings, XmlNode drawingsNode, eChartType? type, ExcelChart topChart, ExcelPivotTable PivotTableSource, XmlDocument chartXml = null, ExcelGroupShape parent = null, string drawingPath = "xdr:graphicFrame", string nvPrPath = "xdr:nvGraphicFramePr/xdr:cNvPr") :
            base(drawings, drawingsNode, chartXml, parent, drawingPath, nvPrPath)
        {
            if (type.HasValue) ChartType = type.Value;
            _topChart = topChart;
            CreateNewChart(drawings, topChart, chartXml, type);

            Init(drawings, _chartNode);

            if (chartXml == null)
            {
                SetTypeProperties();
            }
            else
            {
                ChartType = GetChartType(_chartNode.LocalName);
            }

            InitSeries(this, drawings.NameSpaceManager, _chartNode, PivotTableSource != null);
            if (PivotTableSource != null) SetPivotSource(PivotTableSource);


            if (topChart == null)
                LoadAxis();
            else
            {
                _axis = topChart.Axis;
                if (_axis.Length > 0)
                {
                    XAxis = _axis[0];
                    YAxis = _axis[1];
                }
            }
        }
        internal ExcelChartStandard(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent, string drawingPath = "xdr:graphicFrame", string nvPrPath = "xdr:nvGraphicFramePr/xdr:cNvPr") :
           base(drawings, node, chartXml, parent, drawingPath, nvPrPath)
        {
            UriChart = uriChart;
            Part = part;
            ChartXml = chartXml;
            _chartNode = chartNode;
            InitSeries(this, drawings.NameSpaceManager, _chartNode, PivotTableSource != null);
            InitChartLoad(drawings, chartNode);
            ChartType = GetChartType(chartNode.LocalName);
        }
        internal ExcelChartStandard(ExcelChart topChart, XmlNode chartNode, ExcelGroupShape parent, string drawingPath = "xdr:graphicFrame", string nvPrPath = "xdr:nvGraphicFramePr/xdr:cNvPr") :
            base(topChart, chartNode, parent, drawingPath, nvPrPath)
        {
            UriChart = topChart.UriChart;
            Part = topChart.Part;
            ChartXml = topChart.ChartXml;
            _plotArea = topChart.PlotArea;
            _chartNode = chartNode;
            InitSeries(this, topChart._drawings.NameSpaceManager, _chartNode, false);
            InitChartLoad(topChart._drawings, chartNode);
        }
        private void InitChartLoad(ExcelDrawings drawings, XmlNode chartNode)
        {
            bool isPivot = false;
            Init(drawings, chartNode);
            InitSeries(this, drawings.NameSpaceManager, _chartNode, isPivot);
            LoadAxis();
        }
        internal virtual void InitSeries(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot, List<ExcelChartSerie> list = null)
        {
            Series.Init(chart, ns, node, isPivot, list);
        }
        private void Init(ExcelDrawings drawings, XmlNode chartNode)
        {
            _isChartEx = chartNode.NamespaceURI == ExcelPackage.schemaChartExMain;
            _chartXmlHelper = XmlHelperFactory.Create(drawings.NameSpaceManager, chartNode);
            _chartXmlHelper.AddSchemaNodeOrder(new string[] { "date1904", "lang", "roundedCorners", "AlternateContent", "style", "clrMapOvr", "pivotSource", "protection", "chart", "ofPieType", "title", "pivotFmt", "autoTitleDeleted", "view3D", "floor", "sideWall", "backWall", "plotArea", "wireframe", "barDir", "grouping", "scatterStyle", "radarStyle", "varyColors", "ser", "dLbls", "bubbleScale", "showNegBubbles", "firstSliceAng", "holeSize", "dropLines", "hiLowLines", "upDownBars", "marker", "smooth", "shape", "legend", "plotVisOnly", "dispBlanksAs", "gapWidth", "upBars", "downBars", "showDLblsOverMax", "overlap", "bandFmts", "axId", "spPr", "txPr", "printSettings" }, ExcelDrawing._schemaNodeOrderSpPr);
            WorkSheet = drawings.Worksheet;
        }
        #endregion
        #region "Private functions"
        private void SetTypeProperties()
        {
            /******* Grouping *******/
            if (IsTypeClustered())
            {
                Grouping = eGrouping.Clustered;
            }
            else if (IsTypeStacked())
            {
                Grouping = eGrouping.Stacked;
            }
            else if (
            IsTypePercentStacked())
            {
                Grouping = eGrouping.PercentStacked;
            }

            /***** 3D Perspective *****/
            if (IsType3D())
            {
                View3D.RotY = 20;
                View3D.Perspective = 30;    //Default to 30
                if (IsTypePieDoughnut())
                {
                    View3D.RotX = 30;
                }
                else
                {
                    View3D.RotX = 15;
                }
            }
        }
        private void Init3DProperties()
        {
            Floor = new ExcelChartSurface(this, NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:floor", NameSpaceManager));
            BackWall = new ExcelChartSurface(this, NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:backWall", NameSpaceManager));
            SideWall = new ExcelChartSurface(this, NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:sideWall", NameSpaceManager));
        }
        private void CreateNewChart(ExcelDrawings drawings, ExcelChart topChart, XmlDocument chartXml = null, eChartType? type = null)
        {
            if (topChart == null)
            {
                XmlElement graphFrame = TopNode.OwnerDocument.CreateElement("graphicFrame", ExcelPackage.schemaSheetDrawings);
                graphFrame.SetAttribute("macro", "");
                TopNode.AppendChild(graphFrame);
                graphFrame.InnerXml = string.Format("<xdr:nvGraphicFramePr><xdr:cNvPr id=\"{0}\" name=\"Chart 1\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\" /> <a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\" />   </a:graphicData>  </a:graphic>", _id);
                TopNode.AppendChild(TopNode.OwnerDocument.CreateElement("clientData", ExcelPackage.schemaSheetDrawings));

                var package = drawings.Worksheet._package.ZipPackage;
                UriChart = GetNewUri(package, "/xl/charts/chart{0}.xml");

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
                Part = package.CreatePart(UriChart, ContentTypes.contentTypeChart, _drawings._package.Compression);

                StreamWriter streamChart = new StreamWriter(Part.GetStream(FileMode.Create, FileAccess.Write));
                ChartXml.Save(streamChart);
                streamChart.Close();
                package.Flush();

                var chartRelation = drawings.Part.CreateRelationship(UriHelper.GetRelativeUri(drawings.UriDrawing, UriChart), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/chart");
                graphFrame.SelectSingleNode("a:graphic/a:graphicData/c:chart", NameSpaceManager).Attributes["r:id"].Value = chartRelation.Id;
                package.Flush();
                _chartNode = ChartXml.SelectSingleNode(string.Format("c:chartSpace/c:chart/c:plotArea/{0}", GetChartNodeText()), NameSpaceManager);
            }
            else
            {
                ChartXml = topChart.ChartXml;
                Part = topChart.Part;
                _plotArea = topChart.PlotArea;
                UriChart = topChart.UriChart;
                _axis = topChart._axis;

                XmlNode preNode = _plotArea.ChartTypes[_plotArea.ChartTypes.Count - 1].ChartNode;
                _chartNode = ((XmlDocument)ChartXml).CreateElement(GetChartNodeText(), ExcelPackage.schemaChart);
                preNode.ParentNode.InsertAfter(_chartNode, preNode);
                if (topChart.Axis.Length == 0)
                {
                    AddAxis();
                }
                string serieXML = GetChartSerieStartXml(type.Value, int.Parse(topChart.Axis[0].Id), int.Parse(topChart.Axis[1].Id), topChart.Axis.Length > 2 ? int.Parse(topChart.Axis[2].Id) : -1);
                _chartNode.InnerXml = serieXML;
            }
            GetPositionSize();
            if (IsType3D())
            {
                Init3DProperties();
            }
        }
        private void LoadAxis()
        {
            List<ExcelChartAxis> l = new List<ExcelChartAxis>();
            foreach (XmlNode node in _chartNode.ParentNode.ChildNodes)
            {
                if (node.LocalName.EndsWith("Ax"))
                {
                    ExcelChartAxis ax = new ExcelChartAxisStandard(this, NameSpaceManager, node, "c");
                    l.Add(ax);
                }
            }
            _axis = l.ToArray();

            XmlNodeList nl = _chartNode.SelectNodes("c:axId", NameSpaceManager);
            foreach (XmlNode node in nl)
            {                
                string id = node.Attributes["val"].Value;
                var ix = Array.FindIndex(_axis, x => x.Id == id);
                if(ix>=0)
                {
                    if(XAxis==null)
                    {
                        XAxis = _axis[ix];
                    }
                    else
                    {
                        YAxis = _axis[ix];
                        break;
                    }
                }
            }
        }
        internal virtual eChartType GetChartType(string name)
        {
            switch (name)
            {
                case "area3DChart":
                    if (Grouping == eGrouping.Stacked)
                    {
                        return eChartType.AreaStacked3D;
                    }
                    else if (Grouping == eGrouping.PercentStacked)
                    {
                        return eChartType.AreaStacked1003D;
                    }
                    else
                    {
                        return eChartType.Area3D;
                    }
                case "areaChart":
                    if (Grouping == eGrouping.Stacked)
                    {
                        return eChartType.AreaStacked;
                    }
                    else if (Grouping == eGrouping.PercentStacked)
                    {
                        return eChartType.AreaStacked100;
                    }
                    else
                    {
                        return eChartType.Area;
                    }
                case "doughnutChart":
                    return eChartType.Doughnut;
                case "pie3DChart":
                    return eChartType.Pie3D;
                case "pieChart":
                    return eChartType.Pie;
                case "radarChart":
                    return eChartType.Radar;
                case "scatterChart":
                    return eChartType.XYScatter;
                case "surface3DChart":
                case "surfaceChart":
                    return eChartType.Surface;
                case "stockChart":
                    return eChartType.StockHLC;
                default:
                    return 0;
            }
        }
        #region "Xml init Functions"
        private string ChartStartXml(eChartType type)
        {
            StringBuilder xml = new StringBuilder();
            int axID = 1;
            int xAxID = 2;
            int serAxID = HasThirdAxis() ? 3 : -1;

            xml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            xml.AppendFormat("<c:chartSpace xmlns:c=\"{0}\" xmlns:a=\"{1}\" xmlns:r=\"{2}\">", ExcelPackage.schemaChart, ExcelPackage.schemaDrawings, ExcelPackage.schemaRelationships);
            xml.Append("<c:chart>");
            xml.AppendFormat("{0}{1}<c:plotArea><c:layout/>", AddPerspectiveXml(type), Add3DXml(type));

            string chartNodeText = GetChartNodeText();
            if(type==eChartType.StockVHLC || type==eChartType.StockVOHLC)
            {
                AppendStockChartXml(type, xml, chartNodeText);
            }
            else
            {
                xml.AppendFormat("<{0}>", chartNodeText);
                xml.Append(GetChartSerieStartXml(type, axID, xAxID, serAxID));
                xml.AppendFormat("</{0}>", chartNodeText);
            }

            //Axis
            if (!IsTypePieDoughnut())
            {
                if (IsTypeScatter() || IsTypeBubble())
                {
                    xml.AppendFormat("<c:valAx><c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/>{2}<c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/></c:valAx>", axID, xAxID, GetAxisShapeProperties());
                }
                else
                {
                    xml.AppendFormat("<c:catAx><c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/>{2}<c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/></c:catAx>", axID, xAxID, GetAxisShapeProperties());
                }
                xml.AppendFormat("<c:valAx><c:axId val=\"{1}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:majorGridlines/><c:tickLblPos val=\"nextTo\"/>{2}<c:crossAx val=\"{0}\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/></c:valAx>", axID, xAxID, GetAxisShapeProperties());
                if (serAxID == 3) //Surface Chart
                {
                    if (IsTypeSurface() || ChartType==eChartType.Line3D)
                    {
                        xml.AppendFormat("<c:serAx><c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/>{2}<c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/></c:serAx>", serAxID, xAxID, GetAxisShapeProperties());
                    }
                    else
                    {
                        xml.AppendFormat("<c:valAx><c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"r\"/><c:majorGridlines/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/>{2}<c:crossAx val=\"{1}\"/><c:crosses val=\"max\"/><c:crossBetween val=\"between\"/></c:valAx>", serAxID, axID, GetAxisShapeProperties());
                    }
                }
            }

            xml.AppendFormat("</c:plotArea>" +      //<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>
                AddLegend() +
               "<c:plotVisOnly val=\"1\"/></c:chart>", axID, xAxID);

            xml.Append("<c:printSettings><c:headerFooter/><c:pageMargins b=\"0.75\" l=\"0.7\" r=\"0.7\" t=\"0.75\" header=\"0.3\" footer=\"0.3\"/><c:pageSetup/></c:printSettings></c:chartSpace>");
            return xml.ToString();
        }

        private void AppendStockChartXml(eChartType type, StringBuilder xml, string chartNodeText)
        {
            xml.Append("<c:barChart>");
            xml.Append(AddAxisId(1, 2, -1));
            xml.Append("</c:barChart>");
            xml.AppendFormat("<{0}>", chartNodeText);
            xml.Append(GetChartSerieStartXml(type, 1, 3, -1));
            xml.AppendFormat("</{0}>", chartNodeText);
        }

        private object GetAxisShapeProperties()
        {
            return //"<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>" +
                "<c:txPr>" +
                "<a:bodyPr rot=\"-60000000\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>" +
                "<a:lstStyle/>" +
                "<a:p><a:pPr><a:defRPr kern=\"1200\" sz=\"900\"/></a:pPr></a:p>" +
                "</c:txPr>";
        }

        private string AddLegend()
        {
            return "<c:legend><c:legendPos val=\"r\"/><c:layout/><c:overlay val=\"0\" />" +
                //"<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>" +
                "<c:txPr><a:bodyPr anchorCtr=\"1\" anchor=\"ctr\" wrap=\"square\" vert=\"horz\" vertOverflow=\"ellipsis\" spcFirstLastPara=\"1\" rot=\"0\"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr/></a:p></c:txPr>" +
                "</c:legend>";
        }

        private string GetChartSerieStartXml(eChartType type, int axID, int xAxID, int serAxID)
        {
            StringBuilder xml = new StringBuilder();

            xml.Append(AddScatterType(type));
            xml.Append(AddRadarType(type));
            xml.Append(AddBarDir(type));
            xml.Append(AddGrouping());
            xml.Append(AddVaryColors());
            xml.Append(AddHasMarker(type));
            xml.Append(AddShape(type));
            xml.Append(AddFirstSliceAng(type));
            xml.Append(AddHoleSize(type));
            if (ChartType == eChartType.BarStacked100 ||
                ChartType == eChartType.BarStacked ||
                ChartType == eChartType.ColumnStacked ||
                ChartType == eChartType.ColumnStacked100)
            {
                xml.Append("<c:overlap val=\"100\"/>");
            }
            if (IsTypeSurface())
            {
                xml.Append("<c:bandFmts/>");
            }
            xml.Append(AddAxisId(axID, xAxID, serAxID));

            return xml.ToString();
        }
        private string AddAxisId(int axID, int xAxID, int serAxID)
        {
            if (!IsTypePieDoughnut())
            {
                if (serAxID>0)
                {
                    return string.Format("<c:axId val=\"{0}\"/><c:axId val=\"{1}\"/><c:axId val=\"{2}\"/>", axID, xAxID, serAxID);
                }
                else
                {
                    return string.Format("<c:axId val=\"{0}\"/><c:axId val=\"{1}\"/>", axID, xAxID);
                }
            }
            else
            {
                return "";
            }
        }
        private string AddAxType()
        {
            switch (ChartType)
            {
                case eChartType.XYScatter:
                case eChartType.XYScatterLines:
                case eChartType.XYScatterLinesNoMarkers:
                case eChartType.XYScatterSmooth:
                case eChartType.XYScatterSmoothNoMarkers:
                case eChartType.Bubble:
                case eChartType.Bubble3DEffect:
                    return "valAx";
                default:
                    return "catAx";
            }
        }
        private string AddScatterType(eChartType type)
        {
            if (type == eChartType.XYScatter ||
                type == eChartType.XYScatterLines ||
                type == eChartType.XYScatterLinesNoMarkers ||
                type == eChartType.XYScatterSmooth ||
                type == eChartType.XYScatterSmoothNoMarkers)
            {
                return "<c:scatterStyle val=\"\" />";
            }
            else
            {
                return "";
            }
        }
        private string AddRadarType(eChartType type)
        {
            if (type == eChartType.Radar ||
                type == eChartType.RadarFilled ||
                type == eChartType.RadarMarkers)
            {
                return "<c:radarStyle val=\"\" />";
            }
            else
            {
                return "";
            }
        }
        private string AddGrouping()
        {
            //IsTypeClustered() || IsTypePercentStacked() || IsTypeStacked() || 
            if (IsTypeShape() || IsTypeLine())
            {
                return "<c:grouping val=\"standard\"/>";
            }
            else
            {
                return "";
            }
        }
        private string AddHoleSize(eChartType type)
        {
            if (type == eChartType.Doughnut ||
                type == eChartType.DoughnutExploded)
            {
                return "<c:holeSize val=\"50\" />";
            }
            else
            {
                return "";
            }
        }
        private string AddFirstSliceAng(eChartType type)
        {
            if (type == eChartType.Doughnut ||
                type == eChartType.DoughnutExploded)
            {
                return "<c:firstSliceAng val=\"0\" />";
            }
            else
            {
                return "";
            }
        }
        private string AddVaryColors()
        {
            if (IsTypeStock() || IsTypeSurface())
            {
                return "";
            }
            else
            {
                if (IsTypePieDoughnut())
                {
                    return "<c:varyColors val=\"1\" />";
                }
                else
                {
                    return "<c:varyColors val=\"0\" />";
                }
            }
        }
        private string AddHasMarker(eChartType type)
        {
            if (type == eChartType.LineMarkers ||
                type == eChartType.LineMarkersStacked ||
                type == eChartType.LineMarkersStacked100 /*||
               type == eChartType.XYScatterLines ||
               type == eChartType.XYScatterSmooth*/)
            {
                return "<c:marker val=\"1\"/>";
            }
            else
            {
                return "";
            }
        }
        private string AddShape(eChartType type)
        {
            if (IsTypeShape())
            {
                return "<c:shape val=\"box\" />";
            }
            else
            {
                return "";
            }
        }
        private string AddBarDir(eChartType type)
        {
            if (IsTypeShape())
            {
                return "<c:barDir val=\"col\" />";
            }
            else
            {
                return "";
            }
        }
        private string AddPerspectiveXml(eChartType type)
        {
            //Add for 3D sharts
            if (IsType3D())
            {
                return "<c:view3D><c:perspective val=\"30\" /></c:view3D>";
            }
            else
            {
                return "";
            }
        }
        private string Add3DXml(eChartType type)
        {
            if (IsType3D())
            {
                return Add3DPart("floor") + Add3DPart("sideWall") + Add3DPart("backWall");
            }
            else
            {
                return "";
            }
        }

        private string Add3DPart(string name)
        {
            return string.Format("<c:{0}><c:thickness val=\"0\"/></c:{0}>", name);  //<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/><a:sp3d/></c:spPr>
        }
        #endregion
        #endregion

        /// <summary>
        /// Get the name of the chart node
        /// </summary>
        /// <returns>The name</returns>
        protected string GetChartNodeText()
        {
            switch (ChartType)
            {
                case eChartType.Area3D:
                case eChartType.AreaStacked3D:
                case eChartType.AreaStacked1003D:
                    return "c:area3DChart";
                case eChartType.Area:
                case eChartType.AreaStacked:
                case eChartType.AreaStacked100:
                    return "c:areaChart";
                case eChartType.BarClustered:
                case eChartType.BarStacked:
                case eChartType.BarStacked100:
                case eChartType.ColumnClustered:
                case eChartType.ColumnStacked:
                case eChartType.ColumnStacked100:
                    return "c:barChart";
                case eChartType.Column3D:
                case eChartType.BarClustered3D:
                case eChartType.BarStacked3D:
                case eChartType.BarStacked1003D:
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
                    return "c:bar3DChart";
                case eChartType.Bubble:
                case eChartType.Bubble3DEffect:
                    return "c:bubbleChart";
                case eChartType.Doughnut:
                case eChartType.DoughnutExploded:
                    return "c:doughnutChart";
                case eChartType.Line:
                case eChartType.LineMarkers:
                case eChartType.LineMarkersStacked:
                case eChartType.LineMarkersStacked100:
                case eChartType.LineStacked:
                case eChartType.LineStacked100:
                    return "c:lineChart";
                case eChartType.Line3D:
                    return "c:line3DChart";
                case eChartType.Pie:
                case eChartType.PieExploded:
                    return "c:pieChart";
                case eChartType.BarOfPie:
                case eChartType.PieOfPie:
                    return "c:ofPieChart";
                case eChartType.Pie3D:
                case eChartType.PieExploded3D:
                    return "c:pie3DChart";
                case eChartType.Radar:
                case eChartType.RadarFilled:
                case eChartType.RadarMarkers:
                    return "c:radarChart";
                case eChartType.XYScatter:
                case eChartType.XYScatterLines:
                case eChartType.XYScatterLinesNoMarkers:
                case eChartType.XYScatterSmooth:
                case eChartType.XYScatterSmoothNoMarkers:
                    return "c:scatterChart";
                case eChartType.Surface:
                case eChartType.SurfaceWireframe:
                    return "c:surface3DChart";
                case eChartType.SurfaceTopView:
                case eChartType.SurfaceTopViewWireframe:
                    return "c:surfaceChart";
                case eChartType.StockHLC:
                case eChartType.StockOHLC:
                case eChartType.StockVHLC:
                case eChartType.StockVOHLC:
                    return "c:stockChart";
                default:
                    throw (new NotImplementedException("Chart type not implemented"));
            }
        }
        /// <summary>
        /// Add a secondary axis
        /// </summary>
        internal override void AddAxis()
        {
            XmlElement catAx = ChartXml.CreateElement(string.Format("c:{0}", AddAxType()), ExcelPackage.schemaChart);
            int axID;
            if (_axis.Length == 0)
            {
                _plotArea.TopNode.AppendChild(catAx);
                axID = 1;
            }
            else
            {
                _axis[0].TopNode.ParentNode.InsertAfter(catAx, _axis[_axis.Length - 1].TopNode);
                axID = int.Parse(_axis[0].Id) < int.Parse(_axis[1].Id) ? int.Parse(_axis[1].Id) + 1 : int.Parse(_axis[0].Id) + 1;
            }


            XmlElement valAx = ChartXml.CreateElement("c:valAx", ExcelPackage.schemaChart);
            catAx.ParentNode.InsertAfter(valAx, catAx);

            if (_axis.Length == 0)
            {
                catAx.InnerXml = string.Format("<c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\" /><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/>", axID, axID + 1);
                valAx.InnerXml = string.Format("<c:axId val=\"{1}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\" /><c:axPos val=\"l\"/><c:majorGridlines/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{0}\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/>", axID, axID + 1);
            }
            else
            {
                catAx.InnerXml = string.Format("<c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"1\" /><c:axPos val=\"b\"/><c:tickLblPos val=\"none\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/>", axID, axID + 1);
                valAx.InnerXml = string.Format("<c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\" /><c:axPos val=\"r\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"max\"/><c:crossBetween val=\"between\"/>", axID + 1, axID);
            }

            if (_axis.Length == 0)
            {
                _axis = new ExcelChartAxis[2];
            }
            else
            {
                ExcelChartAxis[] newAxis = new ExcelChartAxis[_axis.Length + 2];
                Array.Copy(_axis, newAxis, _axis.Length);
                _axis = newAxis;
            }

            _axis[_axis.Length - 2] = new ExcelChartAxisStandard(this, NameSpaceManager, catAx, "c");
            _axis[_axis.Length - 1] = new ExcelChartAxisStandard(this, NameSpaceManager, valAx, "c");
            foreach (var chart in _plotArea.ChartTypes)
            {
                chart._axis = _axis;
            }
        }
        internal void RemoveSecondaryAxis()
        {
            throw (new NotImplementedException("Not yet implemented"));
        }
        /// <summary>
        /// Titel of the chart
        /// </summary>
        public new ExcelChartTitle Title
        {
            get
            {
                if (_title == null)
                {
                    _title = new ExcelChartTitle(this, NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart", NameSpaceManager),"c");
                }
                return _title;
            }
        }
        /// <summary>
        /// True if the chart has a title
        /// </summary>
        public override bool HasTitle
        {
            get
            {
                return ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:title", NameSpaceManager) != null;
            }
        }
        /// <summary>
        /// If the chart has a legend
        /// </summary>
        public override bool HasLegend
        {
            get
            {
                return ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:legend", NameSpaceManager) != null;
            }
        }
        /// <summary>
        /// Remove the title from the chart
        /// </summary>
        public override void DeleteTitle()
        {
            _title = null;
            _chartXmlHelper.DeleteNode("../../c:title");
        }
        /// <summary>
        /// The build-in chart styles. 
        /// </summary>
        public override eChartStyle Style
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
                    if (int.TryParse(node.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out int v))
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
                    if (!_chartXmlHelper.ExistsNode("../../../c:style"))
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
        const string _roundedCornersPath = "../../../c:roundedCorners/@val";
        /// <summary>
        /// Border rounded corners
        /// </summary>
        public override bool RoundedCorners
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeBool(_roundedCornersPath);
            }
            set
            {
                _chartXmlHelper.SetXmlNodeBool(_roundedCornersPath, value);
            }
        }
        const string _plotVisibleOnlyPath = "../../c:plotVisOnly/@val";
        /// <summary>
        /// Show data in hidden rows and columns
        /// </summary>
        public override bool ShowHiddenData
        {
            get
            {
                //!!Inverted value!!
                return !_chartXmlHelper.GetXmlNodeBool(_plotVisibleOnlyPath);
            }
            set
            {
                //!!Inverted value!!
                _chartXmlHelper.SetXmlNodeBool(_plotVisibleOnlyPath, !value);
            }
        }
        const string _displayBlanksAsPath = "../../c:dispBlanksAs/@val";
        /// <summary>
        /// Specifies the possible ways to display blanks
        /// </summary>
        public override eDisplayBlanksAs DisplayBlanksAs
        {
            get
            {
                string v = _chartXmlHelper.GetXmlNodeString(_displayBlanksAsPath);
                if (string.IsNullOrEmpty(v))
                {
                    return eDisplayBlanksAs.Zero; //Issue 14715 Changed in Office 2010-?
                }
                else
                {
                    return (eDisplayBlanksAs)Enum.Parse(typeof(eDisplayBlanksAs), v, true);
                }
            }
            set
            {
                _chartXmlHelper.SetXmlNodeString(_displayBlanksAsPath, value.ToString().ToLower(CultureInfo.InvariantCulture));
            }
        }
        const string _showDLblsOverMax = "../../c:showDLblsOverMax/@val";
        /// <summary>
        /// Specifies data labels over the maximum of the chart shall be shown
        /// </summary>
        public override bool ShowDataLabelsOverMaximum
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeBool(_showDLblsOverMax, true);
            }
            set
            {
                _chartXmlHelper.SetXmlNodeBool(_showDLblsOverMax, value, true);
            }
        }
        /// <summary>
        /// Remove all axis that are not used any more
        /// </summary>
        /// <param name="excelChartAxis"></param>
        private void CheckRemoveAxis(ExcelChartAxis excelChartAxis)
        {
            if (ExistsAxis(excelChartAxis))
            {
                //Remove the axis
                ExcelChartAxis[] newAxis = new ExcelChartAxis[Axis.Length - 1];
                int pos = 0;
                foreach (var ax in Axis)
                {
                    if (ax != excelChartAxis)
                    {
                        newAxis[pos] = ax;
                    }
                }

                //Update all charttypes.
                foreach (ExcelChart chartType in _plotArea.ChartTypes)
                {
                    chartType._axis = newAxis;
                }
            }
        }
        private bool ExistsAxis(ExcelChartAxis excelChartAxis)
        {
            foreach (ExcelChart chartType in _plotArea.ChartTypes)
            {
                if (chartType != this)
                {
                    if (chartType.XAxis.AxisPosition == excelChartAxis.AxisPosition ||
                       chartType.YAxis.AxisPosition == excelChartAxis.AxisPosition)
                    {
                        //The axis exists
                        return true;
                    }
                }
            }
            return false;
        }
        /// <summary>
        /// Plotarea
        /// </summary>
        public override ExcelChartPlotArea PlotArea
        {
            get
            {
                if (_plotArea == null)
                {
                    _plotArea = new ExcelChartPlotArea(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:plotArea", NameSpaceManager), this, "c");
                }
                return _plotArea;
            }
        }
        /// <summary>
        /// Legend
        /// </summary>
        public new ExcelChartLegend Legend
        {
            get
            {
                if (_legend == null)
                {
                    _legend = new ExcelChartLegend(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:legend", NameSpaceManager), this, "c");
                }
                return _legend;
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
                    _border = new ExcelDrawingBorder(this, NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace", NameSpaceManager), "c:spPr/a:ln", _chartXmlHelper.SchemaNodeOrder);
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
                    _fill = new ExcelDrawingFill(this, NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace", NameSpaceManager), "c:spPr", _chartXmlHelper.SchemaNodeOrder);
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
                    _effect = new ExcelDrawingEffectStyle(this, NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace", NameSpaceManager), "c:spPr/a:effectLst", _chartXmlHelper.SchemaNodeOrder);
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
                    _threeD = new ExcelDrawing3D(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace", NameSpaceManager), "c:spPr", _chartXmlHelper.SchemaNodeOrder);
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
                    _font = new ExcelTextFont(this, NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace", NameSpaceManager), "c:txPr/a:p/a:pPr/a:defRPr", _chartXmlHelper.SchemaNodeOrder);
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
                    _textBody = new ExcelTextBody(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace", NameSpaceManager), "c:txPr/a:bodyPr", _chartXmlHelper.SchemaNodeOrder);
                }
                return _textBody;
            }
        }

        /// <summary>
        /// 3D-settings
        /// </summary>
        public override ExcelView3D View3D
        {
            get
            {
                if (IsType3D())
                {
                    return new ExcelView3D(NameSpaceManager, ChartXml.SelectSingleNode("//c:view3D", NameSpaceManager));
                }
                else
                {
                    throw (new Exception("Charttype does not support 3D"));
                }

            }
        }
        string _groupingPath = "c:grouping/@val";
        /// <summary>
        /// Specifies the kind of grouping for a column, line, or area chart
        /// </summary>
        public eGrouping Grouping
        {
            get
            {
                return GetGroupingEnum(_chartXmlHelper.GetXmlNodeString(_groupingPath));
            }
            internal set
            {
                _chartXmlHelper.SetXmlNodeString(_groupingPath, GetGroupingText(value));
            }
        }
        string _varyColorsPath = "c:varyColors/@val";
        /// <summary>
        /// If the chart has only one serie this varies the colors for each point.
        /// </summary>
        public override bool VaryColors
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeBool(_varyColorsPath);
            }
            set
            {
                if (value)
                {
                    _chartXmlHelper.SetXmlNodeString(_varyColorsPath, "1");
                }
                else
                {
                    _chartXmlHelper.SetXmlNodeString(_varyColorsPath, "0");
                }
            }
        }

        #region "Grouping Enum Translation"
        private string GetGroupingText(eGrouping grouping)
        {
            switch (grouping)
            {
                case eGrouping.Clustered:
                    return "clustered";
                case eGrouping.Stacked:
                    return "stacked";
                case eGrouping.PercentStacked:
                    return "percentStacked";
                default:
                    return "standard";

            }
        }
        private eGrouping GetGroupingEnum(string grouping)
        {
            switch (grouping)
            {
                case "stacked":
                    return eGrouping.Stacked;
                case "percentStacked":
                    return eGrouping.PercentStacked;
                default: //"clustered":               
                    return eGrouping.Clustered;
            }
        }
        #endregion

        internal int Items
        {
            get
            {
                return 0;
            }
        }

        internal void SetPivotSource(ExcelPivotTable pivotTableSource)
        {
            PivotTableSource = pivotTableSource;
            XmlElement chart = ChartXml.SelectSingleNode("c:chartSpace/c:chart", NameSpaceManager) as XmlElement;

            var pivotSource = ChartXml.CreateElement("pivotSource", ExcelPackage.schemaChart);
            chart.ParentNode.InsertBefore(pivotSource, chart);
            pivotSource.InnerXml = string.Format("<c:name>[]{0}!{1}</c:name><c:fmtId val=\"0\"/>", PivotTableSource.WorkSheet.Name, pivotTableSource.Name);

            var fmts = ChartXml.CreateElement("pivotFmts", ExcelPackage.schemaChart);
            chart.PrependChild(fmts);
            fmts.InnerXml = "<c:pivotFmt><c:idx val=\"0\"/><c:marker><c:symbol val=\"none\"/></c:marker></c:pivotFmt>";

            Series.AddPivotSerie(pivotTableSource);
        }
    }
}
