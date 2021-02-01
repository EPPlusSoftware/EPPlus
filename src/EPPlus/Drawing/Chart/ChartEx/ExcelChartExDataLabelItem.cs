/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/27/2020         EPPlus Software AB           EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// An individual datalabel item
    /// </summary>
    public class ExcelChartExDataLabelItem : ExcelChartExDataLabel
    {
        internal ExcelChartExDataLabelItem(ExcelChartExSerie serie, XmlNamespaceManager nsm, XmlNode node) : base(serie, nsm, node)
        {
        }
        internal ExcelChartExDataLabelItem(ExcelChartExSerie serie, XmlNamespaceManager nsm, XmlNode node, int index) : base(serie, nsm, node)
        {
            Index = index;
        }
        /// <summary>
        /// The index of the datapoint the label is attached to
        /// </summary>
        public int Index 
        { 
            get
            {
                return GetXmlNodeInt("@idx");
            }
            private set
            {                
                SetXmlNodeInt("@idx", value);
            }
        }
    }
}
