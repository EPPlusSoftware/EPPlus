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
using System.Drawing;

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

        internal Color GetRgbColor()
        {
            var h = Hue;
            var s = Saturation / 100;
            var l = Luminance / 100;
            return GetRgb(h, s, l);
        }

        internal static Color GetRgb(double h, double s, double l)
        {
            //Created using formulas here...https://www.rapidtables.com/convert/color/hsl-to-rgb.html
            double r, g, b;

            if (h < 0) h = 0;
            if (s < 0) s = 0;
            if (l < 0) l = 0;
            if (h >= 360) h = 359.99;
            if (s > 1) s = 1;
            if (l > 1) l = 1;

            if (l == 0) return Color.FromArgb(0, 0, 0);
            if (s == 0)
            {
                var c = (int)Math.Round(l * 255,0);
                return Color.FromArgb(c, c, c);
            }
            else
            {
                var c = (1 - Math.Abs(2 * l - 1)) * s;
                var x = c * (1 - Math.Abs((h / 60) % 2 - 1));
                var m = l - c / 2;

                if (h < 60)
                {
                    r = c;
                    g = x;
                    b = 0;
                }
                else if (h < 120)
                {
                    r = x;
                    g = c;
                    b = 0;
                }
                else if (h < 180)
                {
                    r = 0;
                    g = c;
                    b = x;
                }
                else if (h < 240)
                {
                    r = 0;
                    g = x;
                    b = c;
                }
                else if (h < 300)
                {
                    r = x;
                    g = 0;
                    b = c;
                }
                else
                {
                    r = c;
                    g = 0;
                    b = x;
                }
                
                var red = (int)Math.Round(255 * (r + m), 0);
                var green = (int)Math.Round(255 * (g + m), 0);
                var blue = (int)Math.Round(255 * (b + m), 0);

                return Color.FromArgb(red, green, blue);
            }
        }
    }
}