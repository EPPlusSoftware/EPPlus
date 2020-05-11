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
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    public class ExcelChartExDataLabelCollection : ExcelChartExDataLabel, IDrawingStyle, IEnumerable<ExcelChartExDataLabelItem>
    {
        internal ExcelChartExDataLabelCollection(ExcelChartExSerie serie, XmlNamespaceManager ns, XmlNode node, string[] schemaNodeOrder) : 
            base(serie, ns, node)
        {
            _chart = serie._chart;
            AddSchemaNodeOrder(schemaNodeOrder, new string[]{ "numFmt","spPr", "txPr", "visibility", "separator"});
        }
        const string _seriesNameVisiblePath = "cx:visibility/@seriesName";
        public bool SeriesNameVisible
        { 
            get
            {
                return GetXmlNodeBool(_seriesNameVisiblePath);
            }
            set
            {
                SetXmlNodeBool(_seriesNameVisiblePath, value);
            }
        }
        const string _categoryNameVisiblePath = "cx:visibility/@categoryName";
        public bool CategoryNameVisible
        {
            get
            {
                return GetXmlNodeBool(_categoryNameVisiblePath);
            }
            set
            {
                SetXmlNodeBool(_categoryNameVisiblePath, value);
            }
        }
        const string _valueVisiblePath = "cx:visibility/@value";
        public bool ValueVisible
        {
            get
            {
                return GetXmlNodeBool(_valueVisiblePath);
            }
            set
            {
                SetXmlNodeBool(_valueVisiblePath, value);
            }
        }
        const string _separatorPath = "cx:separator";
        //public string Separator 
        //{
        //    get
        //    {
        //        return GetXmlNodeString(_separatorPath);
        //    }
        //    set
        //    {
        //        SetXmlNodeString(_separatorPath, value, true);
        //    }
        //}
        void IDrawingStyleBase.CreatespPr()
        {
            CreatespPrNode("cx:spPr");
        }

        public IEnumerator<ExcelChartExDataLabelItem> GetEnumerator()
        {
            throw new System.NotImplementedException();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new System.NotImplementedException();
        }
    }
}
