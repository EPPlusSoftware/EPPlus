using OfficeOpenXml.Utils.Extensions;
using System;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml
{
    public class ExcelVmlDrawingColor : XmlHelper
    {
        string _path;
        internal ExcelVmlDrawingColor(XmlNamespaceManager ns, XmlNode topNode, string path) : base (ns, topNode)
        {
            _path = path;
        }
        /// <summary>
        /// A color string representing a color. Uses the HTML 4.0 color names, rgb decimal triplets or rgb hex triplets
        /// Example: 
        /// ColorString = "rgb(200,100, 0)"
        /// ColorString = "#FF0000"
        /// ColorString = "Red"
        /// ColorString = "#345" //This is the same as #334455
        /// </summary>
        public string ColorString 
        { 
            get
            {
                return GetXmlNodeString(_path);
            }
            set
            {
                SetXmlNodeString(_path, value);
            }
        }
        /// <summary>
        /// Sets the Color string with the color supplied.
        /// </summary>
        /// <param name="color"></param>
        public void SetColor(Color color)
        {
            ColorString = "#" + (color.ToArgb() & 0xFFFFFF).ToString("X").PadLeft(6, '0');
        }
        /// <summary>
        /// Gets the color for the color string
        /// </summary>
        /// <returns></returns>
        public Color GetColor()
        {
            return GetColor(ColorString);
        }
        internal static Color GetColor(string c)
        {
            if (string.IsNullOrEmpty(c)) return Color.Empty;
            try
            {                
                if (c.IndexOf("[", StringComparison.OrdinalIgnoreCase) > 0)
                {
                    c = c.Substring(0, c.IndexOf("[")).Trim();
                }
                var ts = c.Replace(" ", "");
                if (ts.StartsWith("rgb(",StringComparison.InvariantCultureIgnoreCase))
                {
                    var l = ts.Substring(4, ts.Length - 5).Split(',');
                    if(l.Length==3)
                    {
                        return Color.FromArgb(0xFF, int.Parse(l[0]), int.Parse(l[1]), int.Parse(l[2]));
                    }
                    return Color.Empty;
                }
                else
                {
#if NETSTANDARD
                    return OfficeOpenXml.Compatibility.System.Drawing.ColorTranslator.FromHtml(c);
#else
                    return System.Drawing.ColorTranslator.FromHtml(c);
#endif

                }
            }
            catch
            {
                return Color.Empty;
            }
        }
    }
}