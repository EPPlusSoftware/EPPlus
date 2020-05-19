/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Chart.ChartEx;
<<<<<<< HEAD
using OfficeOpenXml.Drawing.Chart.ChartEx;
=======
>>>>>>> parent of c9b9039... WIP:Added typed classes for Sunburst and treemap charts. More properties and fixed issues.
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Xml;
namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    public class ExcelChartExHistogramSerie : ExcelChartExSerie
    {
        public ExcelChartExHistogramSerie(ExcelChart chart, XmlNamespaceManager ns, XmlNode node) : base(chart, ns, node)
        {

        }
    }
    /// <summary>
    /// A chart serie
    /// </summary>
    public class ExcelChartExSerie : ExcelChartSerie
    {
        XmlNode _dataNode;
        XmlHelper _dataHelper;
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="chart">The chart</param>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="node">Topnode</param>
<<<<<<< HEAD
        internal ExcelChartExSerie(ExcelChartEx chart, XmlNamespaceManager ns, XmlNode node)
            : base(chart, ns, node, "cx")
=======
        internal ExcelChartExSerie(ExcelChart chart, XmlNamespaceManager ns, XmlNode node)
            : base(chart, ns, node)
>>>>>>> parent of c9b9039... WIP:Added typed classes for Sunburst and treemap charts. More properties and fixed issues.
        {
            SchemaNodeOrder = new string[] { "tx", "spPr", "valueColors", "valueColorPositions", "dataPt", "dataLabels", "dataId", "layoutPr", "axisId" };
            _dataNode = node.SelectSingleNode($"../../../../cx:chartData/cx:data[@id={DataId}]", ns);
            _dataHelper = XmlHelperFactory.Create(ns, _dataNode);
            if((chart.ChartType == eChartType.BoxWhisker ||
                chart.ChartType == eChartType.Histogram ||
                chart.ChartType == eChartType.Pareto ||
                chart.ChartType == eChartType.Waterfall ||
                chart.ChartType == eChartType.Pareto) && chart.Series.Count==0)
            {
                AddAxis();
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
                    axis0.InnerXml = "<cx:valScaling/><cx:majorGridlines/><cx:tickLabels/>";
                    axis1.InnerXml = "<cx:catScaling gapWidth=\"1\"/><cx:tickLabels/>";
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
                SetXmlNodeString("cx:tx/cx:txData/cx:f", value.Address);
            }
        }
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

        /// <summary>
        /// Set this to a valid address or the drawing will be invalid.
        /// </summary>
        public override string Series
        {
            get
            {
                return _dataHelper.GetXmlNodeString("*[1]/cx:f");
            }
            set
            {
                _dataHelper.SetXmlNodeString("*[1]/cx:f", value);
            }
        }
        /// <summary>
        /// Set an address for the horizontal labels
        /// </summary>
        public override string XSeries
        {
            get
            {
                return _dataHelper.GetXmlNodeString("*[2]/cx:f");
            }
            set
            {
                _dataHelper.SetXmlNodeString("*[2]/cx:f", value);
            }
        }
        ExcelChartExSerieDataLabel _dataLabels = null;
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
        public eRegionLabelLayout RegionLableLayout 
        {
            get
            {
                return GetXmlNodeString("cx:layoutPr/cx:regionLabelLayout/@val").ToEnum(eRegionLabelLayout.None);
            }
            set
            {
                SetXmlNodeString("cx:layoutPr/cx:regionLabelLayout/@val", value.ToEnumString());
            }
        }
        internal const string _aggregationPath = "cx:layoutPr/cx:aggregation";
        public bool Aggregation
        {
            get
            {
                return ExistNode(_aggregationPath);
            }
            set
            {
                if(value)
                {
                    CreateNode(_aggregationPath);
                }
                else
                {
                    DeleteNode(_aggregationPath);
                }
            }
        }
        ExcelChartExSerieBinning _binning = null;
        /// <summary>
        /// The data binning properties
        /// </summary>
        public ExcelChartExSerieBinning Binning
        {
            get
            {
                if(_binning==null)
                {
                    _binning = new ExcelChartExSerieBinning(NameSpaceManager, TopNode);
                }
                return _binning;
            }
        }
        public ExcelChartExSerieGeography Geography 
        { 
            get; 
            set; 
        }
        public eParentLabelLayout ParentLabelLayout
        {
            get
            {
                return GetXmlNodeString("cx:layoutPr/cx:parentLabelLayout/@val").ToEnum(eParentLabelLayout.None);
            }
            set
            {
                SetXmlNodeString("cx:layoutPr/cx:parentLabelLayout/@val", value.ToEnumString());
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

        public override int NumberOfItems => 0;

        public override ExcelChartTrendlineCollection TrendLines => throw new System.NotImplementedException();

        internal override void SetID(string id)
        {
            throw new System.NotImplementedException();
        }

        internal static XmlElement CreateSeriesAndDataElement(ExcelChartEx chart)
        {
            XmlElement ser = CreateSeriesElement(chart, chart.ChartType, chart.Series.Count);
            ser.InnerXml = $"<cx:dataId val=\"{chart.Series.Count}\"/><cx:layoutPr/>{AddAxisReferense(chart)}";
            SetLayoutProperties(chart, ser);

            chart._chartXmlHelper.CreateNode("../cx:chartData", true);
            var dataElement = (XmlElement)chart._chartXmlHelper.CreateNode("../cx:chartData/cx:data", false, true);
            dataElement.SetAttribute("id", chart.Series.Count.ToString());
            dataElement.InnerXml = $"<cx:strDim type=\"cat\"><cx:f></cx:f></cx:strDim><cx:numDim type=\"{GetNumType(chart.ChartType)}\"><cx:f></cx:f></cx:numDim>";
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
                default:
                    return "val";
            }
        }
    }
}
