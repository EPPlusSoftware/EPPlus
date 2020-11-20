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
                return GetXmlNodeString("v:fill/@type").ToEnum(eVmlFillType.Solid);
            }
            set
            {
                SetXmlNodeString("v:fill/@type", value.ToEnumString());
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
        /// Opacity for fill color 1. Spans 0-100%
        /// </summary>
        public double Opacity
        {
            get
            {
                return ConvertUtil.GetOpacityFromStringVml(GetXmlNodeString("v:fill/@opacity"));
            }
            set
            {
                SetXmlNodeDouble("v:fill/@opacity", value);
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
        public int Recolor { get; set; }
        public int Rotate { get; set; }
    }

    public class ExcelVmlDrawingGradientFill : XmlHelper
    {
        internal ExcelVmlDrawingGradientFill(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {
        }
        /// <summary>
        /// Fill color 2. 
        /// </summary>
        public double SecondColor
        {
            get
            {
                return ConvertUtil.GetOpacityFromStringVml(GetXmlNodeString("v:fill/@opacity2"));
            }
            set
            {
                SetXmlNodeDouble("v:fill/@opacity2", value);
            }
        }

        /// <summary>
        /// Opacity for fill color 2. Spans 0-100%
        /// </summary>
        public double SecondColorOpacity
        {
            get
            {
                return ConvertUtil.GetOpacityFromStringVml(GetXmlNodeString("v:fill/@opacity2"));
            }
            set
            {
                SetXmlNodeDouble("v:fill/@opacity2", value);
            }
        }
        public string ColorsString
        {
            get
            {
                return GetXmlNodeString("v:fill/@colors");
            }
            set
            {
                SetXmlNodeString("v:fill/@colors", value);
            }
        }
        public double? Angle
        {
            get
            {
                return GetXmlNodeDouble("v:fill/@angle");
            }
            set
            {
                SetXmlNodeDouble("v:fill/@angle", value);
            }
        }
        public double Focus
        {
            get
            {
                return GetXmlNodeDouble("v:fill/@focus");
            }
            set
            {
                SetXmlNodeDouble("v:fill/@focus", value);
            }
        }        
    }
}
