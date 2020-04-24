using OfficeOpenXml.Utils.Extentions;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    public class ExcelChartExDataLabel : ExcelChartDataLabel
    {
        internal ExcelChartExDataLabel(ExcelChart chart, XmlNamespaceManager nsm, XmlNode node) : base(chart, nsm, node, "", "cx")
        {

        }

        public override eLabelPosition Position 
        {
            get
            {
                return GetPosEnum(GetXmlNodeString("@pos"));
            }
            set
            {
                SetXmlNodeString("@pos", GetPosText(value));
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
                SetXmlNodeBool(_showValuePath, value);
            }
        }
        private const string _showCategoryPath = "cx:visibility/@showCategory";
        public override bool ShowCategory 
        {
            get
            {
                return GetXmlNodeBool(_showCategoryPath);
            }
            set
            {
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
                SetXmlNodeString("cx:separator", value);
            }
        }
    }
}
