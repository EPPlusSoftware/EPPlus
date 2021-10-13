using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml
{
    /// <summary>
    /// Fill settings for a vml drawing
    /// </summary>
    public class ExcelVmlDrawingFill : XmlHelper
    {
        internal ExcelDrawings _drawings;
        internal ExcelVmlDrawingFill(ExcelDrawings drawings, XmlNamespaceManager ns, XmlNode topNode, string[] schemaNodeOrder) :
            base(ns, topNode)
        {
            SchemaNodeOrder = schemaNodeOrder;
            _drawings = drawings;
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
        ExcelVmlDrawingColor _secondColor;
        /// <summary>
        /// Fill color 2. 
        /// </summary>
        public ExcelVmlDrawingColor SecondColor
        {
            get
            {
                if (_secondColor == null)
                {
                    _secondColor = new ExcelVmlDrawingColor(NameSpaceManager, TopNode, "v:fill/@color2");
                }
                return _secondColor;
            }
        }
        /// <summary>
        /// Opacity for fill color 2. Spans 0-100%
        /// Transparency is is 100-Opacity
        /// </summary>
        public double SecondColorOpacity
        {
            get
            {
                return VmlConvertUtil.GetOpacityFromStringVml(GetXmlNodeString("v:fill/@o:opacity2"));
            }
            set
            {
                if (value < 0 || value > 100)
                {
                    throw (new ArgumentOutOfRangeException("Opacity ranges from 0 to 100%"));
                }
                SetXmlNodeDouble("v:fill/@o:opacity2", value, null, "%");
            }
        }
        ExcelVmlDrawingGradientFill _gradientSettings = null;
        /// <summary>
        /// Gradient specific settings used when <see cref="Style"/> is set to Gradient or GradientRadial.
        /// </summary>
        public ExcelVmlDrawingGradientFill GradientSettings
        {
            get
            {
                if(_gradientSettings==null)
                {
                    _gradientSettings = new ExcelVmlDrawingGradientFill(this, NameSpaceManager, TopNode);
                }
                return _gradientSettings;
            }
        }
        internal ExcelVmlDrawingPictureFill _patternPictureSettings = null;
        /// <summary>
        /// Image and pattern specific settings used when <see cref="Style"/> is set to Pattern, Tile or Frame.
        /// </summary>
        public ExcelVmlDrawingPictureFill PatternPictureSettings
        {
            get
            {
                if(_patternPictureSettings==null)
                {
                    _patternPictureSettings = new ExcelVmlDrawingPictureFill(this, NameSpaceManager, TopNode);
                }
                return _patternPictureSettings;
            }
        }
        /// <summary>
        /// Recolor with picture
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
        /// Rotate fill with shape
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
