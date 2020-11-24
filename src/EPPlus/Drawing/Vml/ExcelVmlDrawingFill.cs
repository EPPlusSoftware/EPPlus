///  <v:fill color2 = "black" recolor="t" rotate="t" focus="100%" type="gradient"/>
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml
{
    public class ExcelVmlDrawingFill : XmlHelper
    {
        internal ExcelVmlDrawingFill(ExcelDrawings drawings, XmlNamespaceManager ns, XmlNode topNode, string[] schemaNodeOrder) :
            base(ns, topNode)
        {
            SchemaNodeOrder = schemaNodeOrder;//new string[] { "fill", "stroke", "shadow", "path", "textbox", "ClientData", "MoveWithCells", "SizeWithCells", "Anchor", "Locked", "AutoFill", "LockText", "TextHAlign", "TextVAlign", "Row", "Column", "Visible" };
        }
        /// <summary>
        /// The type of fill used in the vml drawing
        /// </summary>
        public eVmlFillType Style
        {
            get
            {
                return GetXmlNodeString("v:fill/@type").ToEnum(eVmlFillType.NoFill);
            }
            set
            {
                if (value == eVmlFillType.NoFill)
                {
                    SetXmlNodeString("@filled", "t");
                    DeleteNode("v:fill");
                }
                else
                {
                    DeleteNode("@filled");
                    SetXmlNodeString("v:fill/@type", value.ToEnumString());
                }
            }
        }
        ExcelVmlDrawingColor _fillColor = null;
        /// <summary>
        /// Gradient settings
        /// </summary>
        public ExcelVmlDrawingColor Color
        {
            get
            {
                if (_fillColor == null)
                {
                    _fillColor = new ExcelVmlDrawingColor(NameSpaceManager, TopNode, "@fillcolor");
                }
                return _fillColor;
            }
        }
        /// <summary>
        /// Opacity for fill color 1. Spans 0-100%. 
        /// Transparency is is 100-Opacity
        /// </summary>
        public double Opacity
        {
            get
            {
                return VmlConvertUtil.GetOpacityFromStringVml(GetXmlNodeString("v:fill/@opacity"));
            }
            set
            {
                if(value < 0 || value > 100)
                {
                    throw (new ArgumentOutOfRangeException("Opacity ranges from 0 to 100%"));
                }
                SetXmlNodeDouble("v:fill/@opacity", value, null, "%");
            }
        }
        
        ExcelVmlDrawingGradientFill _gradientSettings = null;
        public ExcelVmlDrawingGradientFill GradientSettings
        {
            get
            {
                if(_gradientSettings==null)
                {
                    _gradientSettings = new ExcelVmlDrawingGradientFill(NameSpaceManager, TopNode);
                }
                return _gradientSettings;
            }
        }
        public bool Recolor 
        { 
            get
            {
                return GetXmlNodeBool("v:fill/@recolor");
            }
            set
            {
                SetXmlNodeBoolVml("v:fill/@recolor", value);
            }
        }
        public bool Rotate 
        {
            get
            {
                return GetXmlNodeBool("v:fill/@rotate");
            }
            set
            {
                SetXmlNodeBoolVml("v:fill/@rotate", value);
            }
        }
    }
}
