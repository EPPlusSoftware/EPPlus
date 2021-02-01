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
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.Style
{
    /// <summary>
    /// A layout the marker of the chart
    /// </summary>
    public class ExcelChartStyleMarkerLayout : XmlHelper
    {

        internal ExcelChartStyleMarkerLayout(XmlNamespaceManager ns, XmlNode topNode) : base(ns, topNode)
        {

        }
        /// <summary>
        /// The marker style
        /// </summary>
        public eMarkerStyle Style
        {
            get
            {
                return GetXmlNodeString("@symbol").ToEnum(eMarkerStyle.None);
            }
            set
            {
                SetXmlNodeString("@symbol", value.ToEnumString());
            }
        }
        /// <summary>
        /// The size of the marker.
        /// Ranges from 2 to 72
        /// </summary>
        public int Size
        {
            get
            {
                var v = GetXmlNodeInt("@size");
                if (v < 0)
                {
                    return 5;   //Default value;
                }
                return v;
            }
            set
            {
                if (value < 2 || value > 72)
                {
                    throw (new ArgumentOutOfRangeException("Marker size must be between 2 and 72"));
                }
                SetXmlNodeString("@size", value.ToString(CultureInfo.InvariantCulture));
            }
        }
    }
}