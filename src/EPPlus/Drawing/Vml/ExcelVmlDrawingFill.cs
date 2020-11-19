///  <v:fill color2 = "black" recolor="t" rotate="t" focus="100%" type="gradient"/>
using OfficeOpenXml.Utils.Extentions;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml
{
    public class ExcelVmlDrawingFill : XmlHelper
    {
        internal ExcelVmlDrawingFill(XmlNode topNode, XmlNamespaceManager ns) :
            base(ns, topNode)
        {
            SchemaNodeOrder = new string[] { "fill", "stroke", "shadow", "path", "textbox", "ClientData", "MoveWithCells", "SizeWithCells", "Anchor", "Locked", "AutoFill", "LockText", "TextHAlign", "TextVAlign", "Row", "Column", "Visible" };
        }
        /// <summary>
        /// The type of fill used in the vml drawing
        /// </summary>
        public eVmlFillType Style 
        {
            get
            {
                return GetXmlNodeString("type").ToEnum(eVmlFillType.Solid);
            }
            set
            {
                SetXmlNodeString("type", value.ToEnumString());
            }
        }
        /// <summary>
        /// Gradient
        /// </summary>
        public ExcelVmlColor Color
        { 
            get; 
            set;
        }
        public ePresetColor PresetFillcolor 
        {
            get;
            set;
        }
        public int Recolor { get; set; }
        public int Rotate { get; set; }
    }
}
