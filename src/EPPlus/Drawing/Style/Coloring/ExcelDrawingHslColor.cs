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
using System.Xml;
using System.Globalization;

namespace OfficeOpenXml.Drawing.Style.Coloring
{
    /// <summary>
    /// Represents a HSL color
    /// </summary>
    public class ExcelDrawingHslColor : XmlHelper
    {
        internal ExcelDrawingHslColor(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {

        }
        /// <summary>
        /// The hue angle in degrees.
        /// Ranges from 0 to 360
        /// </summary>
        public double Hue
        {
            get
            {
                return GetXmlNodeAngel("@hue");
            }
            set
            {
                SetXmlNodeAngel("@hue", value, "Hue");
            }
        }
        /// <summary>
        /// The saturation percentage
        /// </summary>
        public double Saturation
        {
            get
            {
                return GetXmlNodePercentage("@sat") ?? 0;
            }
            set
            {
                SetXmlNodePercentage("@sat", value, false);
            }
        }
        /// <summary>
        /// The luminance percentage
        /// </summary>
        public double Luminance
        {
            get
            {
                return GetXmlNodePercentage("@lum") ?? 0;
            }
            set
            {
                SetXmlNodePercentage("@lum", value, false);
            }
        }

        internal const string NodeName = "a:hslClr";
    }
}