using OfficeOpenXml.Utils.Extensions;
using System.Drawing;
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
    }
}