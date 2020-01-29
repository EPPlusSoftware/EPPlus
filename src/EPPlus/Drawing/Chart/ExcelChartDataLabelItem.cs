/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Represents an individual datalabel
    /// </summary>
    public class ExcelChartDataLabelItem : ExcelChartDataLabel
    {
        internal ExcelChartDataLabelItem(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, string nodeName, string[] schemaNodeOrder)
           : base(chart, ns, node, nodeName, schemaNodeOrder)
        {
            
        }
        /// <summary>
        /// The index of an individual datalabel
        /// </summary>
        public int Index
        {
            get
            {
                return GetXmlNodeInt("c:idx/@val");
            }
            set
            {
                SetXmlNodeString("c:idx/@val", value.ToString(CultureInfo.InvariantCulture));
            }
        }
    }
}
