///  <v:fill color2 = "black" recolor="t" rotate="t" focus="100%" type="gradient"/>
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml
{
    public class ExcelVmlDrawingFill : XmlHelper
    {
        internal ExcelVmlDrawingFill(ExcelDrawings drawings, XmlNamespaceManager ns, XmlNode topNode, string[] schemaNodeOrder) :
            base(ns, topNode)
        {
            SchemaNodeOrder = schemaNodeOrder;
            SetSettings(Style);
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
                    SetSettings(value);
                }
            }
        }

        private void SetSettings(eVmlFillType value)
        {
            if(value==eVmlFillType.NoFill || value==eVmlFillType.Solid)
            {
                _gradientSettings = null;
                _patternPictureSettings = null;
            }
            if (_gradientSettings == null && (value == eVmlFillType.Gradient || value == eVmlFillType.GradientRadial))
            {
                _gradientSettings = new ExcelVmlDrawingGradientFill(this, NameSpaceManager, TopNode);
                _patternPictureSettings = null;
            }
            else if (_patternPictureSettings == null && (value == eVmlFillType.Gradient || value == eVmlFillType.GradientRadial))
            {
                _gradientSettings = null;
                _patternPictureSettings = new ExcelVmlDrawingBlipFill(this, NameSpaceManager, TopNode);
            }
        }

        ExcelVmlDrawingColor _fillColor = null;
        /// <summary>
        /// The primary color used for filling the drawing.
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
        /// <summary>
        /// If <see cref="Style"/> is set to Gradient or GradientRadial this propery contains the gradient specific settings.
        /// </summary>
        public ExcelVmlDrawingGradientFill GradientSettings
        {
            get
            {
                return _gradientSettings;
            }
        }
        ExcelVmlDrawingBlipFill _patternPictureSettings = null;
        /// <summary>
        /// If <see cref="Style"/> is set to Gradient or GradientRadial this propery contains the gradient specific settings.
        /// </summary>
        public ExcelVmlDrawingBlipFill PatternPictureSettings
        {
            get
            {
                return _patternPictureSettings;
            }
        }

        /// <summary>
        /// Recolor
        /// </summary>
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
        /// <summary>
        /// Rotate
        /// </summary>
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
