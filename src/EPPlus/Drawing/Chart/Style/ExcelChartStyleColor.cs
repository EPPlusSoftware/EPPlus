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
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.Style
{
    /// <summary>
    /// A color for a chart style entry reference
    /// </summary>
    public class ExcelChartStyleColor : XmlHelper
    {
        internal ExcelChartStyleColor(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {

        }
        /// <summary>
        /// Color is automatic
        /// </summary>
        public bool Auto
        {
            get
            {
                var v = GetXmlNodeString("@val");
                return v == "auto";                    
            }
        }
        /// <summary>
        /// The index, maps to the style matrix in the theme
        /// </summary>
        public int? Index
        {
            get
            {
                return GetXmlNodeIntNull("@val");
            }
        }
        internal void SetValue(bool isAuto, int index)
        {
            if(Auto)
            {
                SetXmlNodeString("@val", "auto");
            }
            else
            {
                if (index < 0) throw new ArgumentOutOfRangeException("index", "Index can't be negative");
                SetXmlNodeString("@val", index.ToString(CultureInfo.InvariantCulture));
            }

        }
    }   
}