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
using OfficeOpenXml.Utils.Extentions;
using System.Xml;
namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    public class ExcelBoxWhiskerChartSerie : ExcelChartExSerie
    {
        public ExcelBoxWhiskerChartSerie(ExcelChartEx chart, XmlNamespaceManager ns, XmlNode node) : base(chart, ns, node)
        {

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
        /// The quartile calculation methods
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
        ExcelChartExSerieElementVisibilities _elementVisibility = null;
        public ExcelChartExSerieElementVisibilities ElementVisibility
        {
            get
            {
                if (_elementVisibility == null)
                {
                    _elementVisibility = new ExcelChartExSerieElementVisibilities(NameSpaceManager, TopNode, SchemaNodeOrder);
                }
                return _elementVisibility;
            }
        }
    }
}
