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
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Xml;
namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
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
        internal ExcelChartExSerie(ExcelChart chart, XmlNamespaceManager ns, XmlNode node)
            : base(chart, ns, node)
        {
            SchemaNodeOrder = new string[] { "tx", "spPr", "valueColors", "valueColorPositions", "dataPt", "dataLabels", "dataId", "layoutPr", "axisId" };
            _dataNode = node.SelectSingleNode($"../../../../cx:chartData/cx:data[@id={DataId}]", ns);
            _dataHelper = XmlHelperFactory.Create(ns, _dataNode);
            _seriesXPath = "cx:strDim";
            _seriesPath = "cx:numDim";
            //foreach (XmlElement e in _dataNode.ChildNodes)
            //{
            //    var t = e.GetAttribute("type");
            //    if(e.LocalName == "numDim" || t!="x")
            //    {
            //        _seriesPath = "cx:numDim";
            //    }
            //    else if(e.LocalName=="strDim")
            //    {
            //        _seriesXPath = "cx:numDim";
            //    }
            //}
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

        string _seriesPath;
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
        string _seriesXPath;
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
                    _dataPoints = new ExcelChartExDataPointCollection(_chart,NameSpaceManager, TopNode, SchemaNodeOrder);
                }
                return _dataPoints;
            }
        }
        ExcelChartExSerieElementVisibilities _elementVisibility = null;
        public ExcelChartExSerieElementVisibilities ElementVisibility
        {
            get
            {
                if(_elementVisibility==null)
                {
                    _elementVisibility = new ExcelChartExSerieElementVisibilities(NameSpaceManager, TopNode, SchemaNodeOrder);
                }
                return _elementVisibility;
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
        /// 
        /// </summary>
        public eQuartileMethod? QuartileMethod 
        { 
            get
            {
                var s = GetXmlNodeString("cx:layoutPr/cx:statistics/@quartileMethod");
                if (string.IsNullOrEmpty(s)) return null;                
                return s.ToEnum(eQuartileMethod.Inclusive);
            }
            set
            {
                SetXmlNodeString("cx:layoutPr/cx:statistics/@quartileMethod", value.ToEnumString());
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

        internal static XmlElement CreateSerieElement(XmlNamespaceManager ns, XmlNode node, ExcelChart chart)
        {
            throw new NotImplementedException();
        }
    }
}
