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
using System.Drawing;
using System.Globalization;
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
        
    }
}
