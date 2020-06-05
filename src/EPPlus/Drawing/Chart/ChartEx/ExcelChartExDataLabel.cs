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
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// Datalabel on chart level. 
    /// </summary>
    public class ExcelChartExDataLabel : ExcelChartDataLabel
    {
        protected internal readonly ExcelChartExSerie _serie;
        internal ExcelChartExDataLabel(ExcelChartExSerie serie, XmlNamespaceManager nsm, XmlNode node) : base(serie._chart, nsm, node, "", "cx")
        {
            _serie = serie;
            SchemaNodeOrder = new string[] { "numFmt","visibility", "spPr","txPr","visibility", "separator" };
        }
        internal const string _dataLabelPath = "cx:dataLabel";
        /// <summary>
        /// The datalabel position
        /// </summary>
        public override eLabelPosition Position 
        {
            get
            {
                return GetPosEnum(GetXmlNodeString("@pos"));
            }
            set
            {
                SetDataLabelNode();
                SetXmlNodeString("@pos", GetPosText(value));
            }
        }

        internal virtual void SetDataLabelNode()
        {
            if (TopNode.LocalName == "series")
            {
                TopNode = _serie.CreateNode("cx:dataLabels");
            }
        }

        private const string _showValuePath = "cx:visibility/@value";
        /// <summary>
        /// Show values in the datalabels
        /// </summary>
        public override bool ShowValue 
        { 
            get
            {
                return GetXmlNodeBool(_showValuePath);
            }
            set
            {
                SetDataLabelNode();
                SetXmlNodeBool(_showValuePath, value);
            }
        }
        private const string _showCategoryPath = "cx:visibility/@categoryName";
        /// <summary>
        /// Show category names in the datalabels
        /// </summary>
        public override bool ShowCategory 
        {
            get
            {
                return GetXmlNodeBool(_showCategoryPath);
            }
            set
            {
                SetDataLabelNode();
                SetXmlNodeBool(_showCategoryPath, value);
            }
        }
        private const string _seriesNamePath = "cx:visibility/@seriesName";
        /// <summary>
        /// Show series names in the datalabels
        /// </summary>
        public override bool ShowSeriesName 
        {
            get
            {
                return GetXmlNodeBool(_seriesNamePath);
            }
            set
            {
                SetDataLabelNode();
                SetXmlNodeBool(_seriesNamePath, value);
            }
        }
        /// <summary>
        /// This property is not used for extended charts. Trying to set this property will result in a NotSupportedException.
        /// </summary>
        public override bool ShowPercent 
        {
            get
            {
                return false;
            }
            set
            {
                throw (new NotSupportedException("ShowPercent do not apply to Extension Charts"));
            }
        }
        /// <summary>
        /// This property is not used for extended charts. Trying to set this property will result in a NotSupportedException.
        /// </summary>
        public override bool ShowLeaderLines 
        {
            get
            {
                return false;
            }
            set
            {
                throw (new NotSupportedException("ShowLeaderLines do not apply to Extension Charts"));
            }
        }
        /// <summary>
        /// This property is not used for extended charts. Trying to set this property will result in a NotSupportedException.
        /// </summary>
        public override bool ShowBubbleSize 
        {
            get
            {
                return false;
            }
            set
            {
                throw (new NotSupportedException("ShowBubbleSize do not apply to Extension Charts"));
            }
        }
        /// <summary>
        /// This property is not used for extended charts. Trying to set this property will result in a NotSupportedException.
        /// </summary>
        public override bool ShowLegendKey 
        {
            get
            {
                return false;
            }
            set
            {
                throw (new InvalidOperationException("ShowLegendKey do not apply to Extension Charts"));
            }
        }
        /// <summary>
        /// The separator between items in the datalabel
        /// </summary>
        public override string Separator
        {
            get
            {
                return GetXmlNodeString("cx:separator");
            }
            set
            {
                SetDataLabelNode();
                SetXmlNodeString("cx:separator", value);
            }
        }
    }
}
