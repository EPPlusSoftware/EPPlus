/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB           Release EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Globalization;
using System.Xml;
namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// A chart serie
    /// </summary>
    public class ExcelChartExSerie : ExcelChartSerie
    {
        XmlNode _dataNode;
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="chart">The chart</param>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="node">Topnode</param>
        internal ExcelChartExSerie(ExcelChartEx chart, XmlNamespaceManager ns, XmlNode node)
            : base(chart, ns, node, "cx")
        {
            SchemaNodeOrder = new string[] { "tx", "spPr", "valueColors", "valueColorPositions", "dataPt", "dataLabels", "dataId", "layoutPr", "axisId" };
            _dataNode = node.SelectSingleNode($"../../../../cx:chartData/cx:data[@id={DataId}]", ns);
            if((chart.ChartType == eChartType.BoxWhisker ||
                chart.ChartType == eChartType.Histogram ||
                chart.ChartType == eChartType.Pareto ||
                chart.ChartType == eChartType.Waterfall ||
                chart.ChartType == eChartType.Pareto) && chart.Series.Count==0)
            {
                if(chart._chartXmlHelper.ExistsNode("cx:plotArea/cx:axis")==false)
                {
                    AddAxis();
                }
                chart.LoadAxis();
            }
        }

        private void AddAxis()
        {
            var axis0=(XmlElement)_chart._chartXmlHelper.CreateNode("cx:plotArea/cx:axis");
            axis0.SetAttribute("id", "0");
            var axis1 = (XmlElement)_chart._chartXmlHelper.CreateNode("cx:plotArea/cx:axis", false, true);
            axis1.SetAttribute("id", "1");

            switch(_chart.ChartType)
            {
                case eChartType.BoxWhisker:
                    axis0.InnerXml = "<cx:catScaling gapWidth=\"1\"/><cx:tickLabels/>";
                    axis1.InnerXml = "<cx:valScaling/><cx:majorGridlines/><cx:tickLabels/>";
                    break;
                case eChartType.Waterfall:
                    axis0.InnerXml = "<cx:catScaling/><cx:tickLabels/>";
                    axis1.InnerXml = "<cx:valScaling/><cx:tickLabels/>";
                    break;
                case eChartType.Funnel:
                    axis1.InnerXml = "<cx:catScaling gapWidth=\"0.06\"/><cx:tickLabels/>";
                    break;
                case eChartType.Histogram:
                case eChartType.Pareto:
                    axis0.InnerXml = "<cx:catScaling gapWidth=\"0\"/><cx:tickLabels/>";
                    axis1.InnerXml = "<cx:valScaling/><cx:majorGridlines/><cx:tickLabels/>";
                    if(_chart.ChartType== eChartType.Pareto)
                    {
                        var axis2 = (XmlElement)_chart._chartXmlHelper.CreateNode("cx:plotArea/cx:axis", false, true);
                        axis2.SetAttribute("id", "2");
                        axis2.InnerXml = "<cx:valScaling min=\"0\" max=\"1\"/><cx:units unit=\"percentage\"/><cx:tickLabels/>";
                    }
                    break;
            }
        }
        internal int DataId
        {
            get
            {
                return GetXmlNodeInt("cx:dataId/@val");
            }
        }
        ExcelChartExDataCollection _dataDimensions = null;
        /// <summary>
        /// The dimensions of the serie
        /// </summary>
        public ExcelChartExDataCollection DataDimensions
        {
            get
            {
                if (_dataDimensions == null)
                {
                    _dataDimensions = new ExcelChartExDataCollection(this, NameSpaceManager, _dataNode);
                }
                return _dataDimensions;
            }
        }
        const string headerAddressPath = "c:tx/c:strRef/c:f";
        /// <summary>
        /// Header address for the serie.
        /// </summary>
        public override ExcelAddressBase HeaderAddress
        {
            get
            {
                var f = GetXmlNodeString("cx:tx/cx:txData/cx:f");
                if (ExcelAddress.IsValidAddress(f))
                {
                    return new ExcelAddressBase(f);
                }
                else
                {
                    if (_chart.WorkSheet.Workbook.Names.ContainsKey(f))
                    {
                        return _chart.WorkSheet.Workbook.Names[f];
                    }
                    else if (_chart.WorkSheet.Names.ContainsKey(f))
                    {
                        return _chart.WorkSheet.Names[f];
                    }
                    return null;
                }
            }
            set
            {
                SetXmlNodeString("cx:tx/cx:txData/cx:f", value.FullAddress);
            }
        }
        /// <summary>
        /// The header text for the serie.
        /// </summary>
        public override string Header
        {
            get
            {
                return GetXmlNodeString("cx:tx/cx:txData/cx:v");
            }
            set
            {
                SetXmlNodeString("cx:tx/cx:txData/cx:v", value);
            }
        }
        XmlHelper _catSerieHelper = null;
        XmlHelper _valSerieHelper = null;
        /// <summary>
        /// Set this to a valid address or the drawing will be invalid.
        /// </summary>
        public override string Series
        {
            get
            {
                var helper = GetSerieHelper();
                return helper.GetXmlNodeString("cx:f");
            }
            set
            {
                var helper = GetSerieHelper();
                helper.SetXmlNodeString("cx:f", ToFullAddress(value));
            }
        }

        /// <summary>
        /// Set an address for the horizontal labels
        /// </summary>
        public override string XSeries
        {
            get
            {
                var helper = GetXSerieHelper(false);
                if(helper==null)
                {
                    return "";
                }
                else
                {
                    return helper.GetXmlNodeString("cx:f");
                }
            }
            set
            {
                var helper = GetXSerieHelper(true);
                helper.SetXmlNodeString("cx:f", ToFullAddress(value));
            }
        }
        private XmlHelper GetSerieHelper()
        {
            if (_valSerieHelper == null)
            {
                if (_dataNode.ChildNodes.Count == 1)
                {
                    _valSerieHelper = XmlHelperFactory.Create(NameSpaceManager, _dataNode.FirstChild);
                }
                else if (_dataNode.ChildNodes.Count > 1)
                {
                    _valSerieHelper = XmlHelperFactory.Create(NameSpaceManager, _dataNode.ChildNodes[1]); 
                }
            }
            return _valSerieHelper;
        }

        private XmlHelper GetXSerieHelper(bool create)
        {
            if (_catSerieHelper == null)
            {
                if (_dataNode.ChildNodes.Count == 1)
                {
                    if (create)
                    {
                        var node = _dataNode.OwnerDocument.CreateElement("cx", "strDim", ExcelPackage.schemaChartExMain);
                        _dataNode.InsertBefore(node, _dataNode.FirstChild);
                        _catSerieHelper = XmlHelperFactory.Create(NameSpaceManager, node);
                    }
                    else
                    {
                        return null;
                    }
                }
                else if (_dataNode.ChildNodes.Count > 1)
                {
                    _catSerieHelper = XmlHelperFactory.Create(NameSpaceManager, _dataNode.ChildNodes[0]); 
                }
            }
            return _catSerieHelper;
        }

        ExcelChartExSerieDataLabel _dataLabels = null;
        /// <summary>
        /// Data label properties
        /// </summary>
        public ExcelChartExSerieDataLabel DataLabel
        {
            get
            {
                if (_dataLabels == null)
                {
                    _dataLabels = new ExcelChartExSerieDataLabel(this, NameSpaceManager, TopNode, SchemaNodeOrder);
                }
                return _dataLabels;
            }
        }
        ExcelChartExDataPointCollection _dataPoints = null;
        /// <summary>
        /// A collection of individual data points
        /// </summary>
        public ExcelChartExDataPointCollection DataPoints
        {
            get
            {
                if(_dataPoints==null)
                {
                    _dataPoints = new ExcelChartExDataPointCollection(this, NameSpaceManager, TopNode, SchemaNodeOrder);
                }
                return _dataPoints;
            }
        }
        /// <summary>
        /// If the serie is hidden
        /// </summary>
        public bool Hidden
        {
            get
            {
                return GetXmlNodeBool("@hidden", false);
            }
            set
            {
                SetXmlNodeBool("@hidden", value, false);
            }
        }

        /// <summary>
        /// If the chart has datalabel
        /// </summary>
        public bool HasDataLabel
        {
            get
            {
                return TopNode.SelectSingleNode("c:dataLabels", NameSpaceManager) != null;
            }
        }

        /// <summary>
        /// Number of items. Will always return 0, as no item data is stored.
        /// </summary>
        public override int NumberOfItems => 0;

        /// <summary>
        /// Trendline do not apply to extended charts.
        /// </summary>
        public override ExcelChartTrendlineCollection TrendLines => new ExcelChartTrendlineCollection(null);

        internal override void SetID(string id)
        {
            throw new System.NotImplementedException();
        }

        internal static XmlElement CreateSeriesAndDataElement(ExcelChartEx chart, bool hasCatSerie)
        {
            XmlElement ser = CreateSeriesElement(chart, chart.ChartType, chart.Series.Count);
            ser.InnerXml = $"<cx:dataId val=\"{chart.Series.Count}\"/><cx:layoutPr/>{AddAxisReferense(chart)}";
            SetLayoutProperties(chart, ser);

            chart._chartXmlHelper.CreateNode("../cx:chartData", true);
            var dataElement = (XmlElement)chart._chartXmlHelper.CreateNode("../cx:chartData/cx:data", false, true);
            dataElement.SetAttribute("id", chart.Series.Count.ToString());
            var innerXml="";
            if (hasCatSerie == true)
            {
                innerXml += $"<cx:strDim type=\"cat\"><cx:f></cx:f><cx:nf></cx:nf></cx:strDim>";
            }
            innerXml += $"<cx:numDim type=\"{GetNumType(chart.ChartType)}\"><cx:f></cx:f><cx:nf></cx:nf></cx:numDim>";
            dataElement.InnerXml = innerXml;
            return ser;
        }

        internal static XmlElement CreateSeriesElement(ExcelChartEx chart, eChartType type, int index, XmlNode referenceNode = null, bool isPareto=false)
        {
            var plotareaNode = chart._chartXmlHelper.CreateNode("cx:plotArea/cx:plotAreaRegion");
            var ser = plotareaNode.OwnerDocument.CreateElement("cx", "series", ExcelPackage.schemaChartExMain);
            XmlNodeList node = plotareaNode.SelectNodes("cx:series", chart.NameSpaceManager);

            if(node.Count > 0)
            {
                plotareaNode.InsertAfter(ser, referenceNode ?? node[node.Count - 1]);
            }
            else
            {
                var f = XmlHelperFactory.Create(chart.NameSpaceManager, plotareaNode);
                f.InserAfter(plotareaNode, "cx:plotSurface", ser);
            }
            ser.SetAttribute("formatIdx", index.ToString());
            ser.SetAttribute("uniqueId", "{" + Guid.NewGuid().ToString() + "}");
            ser.SetAttribute("layoutId", GetLayoutId(type, isPareto));
            return ser;
        }

        private static object AddAxisReferense(ExcelChartEx chart)
        {
            if(chart.ChartType==eChartType.Pareto)
            {
                return "<cx:axisId val=\"1\"/>";
            }
            else
            {
                return "";
            }            
        }

        private static string GetLayoutId(eChartType chartType, bool isPareto)
        {
            if (isPareto) return "paretoLine";
            switch(chartType)
            {
                case eChartType.Histogram:
                case eChartType.Pareto:
                    return "clusteredColumn";
                default:
                    return chartType.ToEnumString();
            }            
        }

        private static void SetLayoutProperties(ExcelChartEx chart, XmlElement ser)
        {
            var layoutPr = ser.SelectSingleNode("cx:layoutPr", chart.NameSpaceManager);
            switch (chart.ChartType)
            {
                case eChartType.BoxWhisker:
                    layoutPr.InnerXml = "<cx:parentLabelLayout val=\"banner\"/><cx:visibility outliers=\"1\" nonoutliers=\"0\" meanMarker=\"1\" meanLine=\"0\"/><cx:statistics quartileMethod=\"exclusive\"/>";
                    break;
                case eChartType.Histogram:
                case eChartType.Pareto:
                    layoutPr.InnerXml = "<cx:binning intervalClosed=\"r\"/>";
                    break;
                case eChartType.RegionMap:
                    var ci = CultureInfo.CurrentCulture;
                    layoutPr.InnerXml = $"<cx:geography attribution = \"Powered by Bing\" cultureRegion = \"{ci.TwoLetterISOLanguageName}\" cultureLanguage = \"{ci.Name}\" ><cx:geoCache provider=\"{{E9337A44-BEBE-4D9F-B70C-5C5E7DAFC167}}\"><cx:binary/></cx:geoCache></cx:geography>";
                    break;
                case eChartType.Waterfall:
                    layoutPr.InnerXml = "<cx:visibility connectorLines=\"0\" />";
                    break;
            }
        }

        private static string GetNumType(eChartType chartType)
        {
            switch (chartType)
            {
                case eChartType.Sunburst:
                case eChartType.Treemap:
                    return "size";
                case eChartType.RegionMap:
                    return "colorVal";
                default:
                    return "val";
            }
        }
    }
}
