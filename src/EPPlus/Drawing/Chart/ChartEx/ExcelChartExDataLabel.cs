/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/15/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    public class ExcelChartExDataLabel : ExcelChartDataLabel
    {
        protected internal readonly ExcelChartExSerie _serie;
        internal ExcelChartExDataLabel(ExcelChartExSerie serie, XmlNamespaceManager nsm, XmlNode node) : base(serie._chart, nsm, node, "", "cx")
        {
            _serie = serie;
        }

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

        private void SetDataLabelNode()
        {
            if (TopNode.LocalName == "series")
            {
                TopNode = _serie.CreateNode("cx:dataLabels");
            }
        }

        private const string _showValuePath = "cx:visibility/@value";
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
        public override bool ShowPercent 
        {
            get;
            set; 
        }
        public override bool ShowLeaderLines 
        {
            get;
            set;
        }
        public override bool ShowBubbleSize 
        {
            get;
            set;
        }
        public override bool ShowLegendKey 
        {
            get;
            set;
        }
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
