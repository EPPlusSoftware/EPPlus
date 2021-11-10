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
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Coloring
{
    /// <summary>
    /// Represents a RGB color
    /// </summary>
    public class ExcelDrawingRgbColor : XmlHelper
    {
        internal ExcelDrawingRgbColor(XmlNamespaceManager nsm, XmlNode topNode) : base (nsm, topNode)
        {
        }
        /// <summary>
        /// The color
        /// </summary>s
        public Color Color
        {
            get
            {
                var s = GetXmlNodeString("@val");
                return GetColorFromString(s);
            }
            set
            {               
                SetXmlNodeString("@val", (value.ToArgb() & 0xFFFFFF).ToString("X").PadLeft(6, '0'));
            }
        }
        internal static Color GetColorFromString(string s)
        {
            int n;
            if (s.Length == 6) s = "FF" + s;
            if (int.TryParse(s, System.Globalization.NumberStyles.HexNumber, CultureInfo.InvariantCulture, out n))
            {
                return Color.FromArgb(n);
            }
            else
            {
                return Color.Empty;
            }
        }

        internal const string NodeName = "a:srgbClr";
        internal void SetXml(XmlNamespaceManager nsm, XmlNode node, bool doInit = false)
        {
            
        }
        internal void GetXml()
        {
        }
        internal void GetHsl(out double hue, out double saturation, out double luminance)
        {
            GetHslColor(Color.R, Color.G, Color.B, out hue, out saturation, out luminance);
        }

        internal static void GetHslColor(Color c, out double hue, out double saturation, out double luminance)
        {
            GetHslColor(c.R, c.G, c.B, out hue, out saturation, out luminance);
        }
        internal static void GetHslColor(byte red, byte green, byte blue, out double hue, out double saturation, out double luminance)
        {
            //Created using formulas here...https://www.rapidtables.com/convert/color/rgb-to-hsl.html
            var r = red / 255D;
            var g = green / 255D;
            var b = blue / 255D;

            var ix = new double[]{ r, g, b };
            var cMax = ix.Max();
            var cMin = ix.Min();
            var delta = cMax - cMin;


            if (delta == 0)
            {
                hue = 0;
            }
            else if (cMax == r)
            {
                hue = 60 * (((g - b) / delta) % 6);
            }
            else if (cMax == g)
            {
                hue = 60 * ((b - r) / delta + 2);
            }
            else
            {
                hue = 60 * ((r - g) / delta + 4);
            }
           
            if (hue < 0)
                hue += 360;

            luminance = (cMax + cMin) / 2;
            saturation = delta == 0 ? 0 : delta / (1 - Math.Abs(2 * luminance - 1));
        }
    }
}
