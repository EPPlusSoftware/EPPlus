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
using OfficeOpenXml.Drawing.Chart.ChartEx.enums;
using OfficeOpenXml.Utils.Extentions;
using System.Xml;
namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// A chart serie
    /// </summary>
    public class ExcelChartExSerie : ExcelChartSerieBase
    {
        XmlNode _dataNode;
        XmlHelper _dataHelper;
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="chart">The chart</param>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="node">Topnode</param>
        internal ExcelChartExSerie(ExcelChartBase chart, XmlNamespaceManager ns, XmlNode node)
            : base(chart, ns, node)
        {
            SchemaNodeOrder = new string[] { "tx", "spPr", "valueColors", "valueColorPositions", "dataPt", "dataLabels", "dataId", "layoutPr", "axisId" };
            _dataNode = node.SelectSingleNode($"../../../../cx:chartData/cx:data[@id={DataId}]", ns);
            _dataHelper = XmlHelperFactory.Create(ns, _dataNode);
        }
        internal int DataId
        {
            get
            {
                return GetXmlNodeInt("cx:dataId/@val");
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
                return _dataHelper.GetXmlNodeString("cx:numDim[@type='val']|cx:strDim[@type='val']");
            }
            set
            {
                _dataHelper.SetXmlNodeString("cx:numDim[@type='val']|cx:strDim[@type='val']", value);
            }
        }
        /// <summary>
        /// Set an address for the horizontal labels
        /// </summary>
        public override string XSeries
        {
            get
            {
                return _dataHelper.GetXmlNodeString("cx:numDim[@type='cat']|cx:strDim[@type='cat']");
            }
            set
            {
                _dataHelper.SetXmlNodeString("cx:numDim[@type='cat']|cx:strDim[@type='cat']", value);
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
    }
}
