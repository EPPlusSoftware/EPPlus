///  <v:fill color2 = "black" recolor="t" rotate="t" focus="100%" type="gradient"/>
using System;
using System.Xml;
using OfficeOpenXml.Utils.Extensions;
namespace OfficeOpenXml.Drawing.Vml
{
    public class ExcelVmlDrawingGradientFill : XmlHelper
    {
        ExcelVmlDrawingFill _fill;
        internal ExcelVmlDrawingGradientFill(ExcelVmlDrawingFill fill, XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {
            _fill = fill;
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
        
        /// <summary>
        /// A list of colors used for gradient fill. 
        /// Each color item starts with a percent and a color. Starting from 0% and endig and 100%
        /// <see cref="SetGradientColors"></see>
        /// </summary>
        public string ColorsString
        {
            get
            {
                return GetXmlNodeString("v:fill/@colors");
            }
            internal set
            {
                SetXmlNodeString("v:fill/@colors", value);
            }
        }
        /// <summary>
        /// Sets the <see cref="ColorsString"/> with the colors supplied
        /// Each color item starts with a percent and a color. 
        /// Percent values must be sorted, starting from 0% and ending and 100%.
        /// </summary>
        /// <param name="colors">The colors with a percent value for the gradient fill</param>
        public void SetGradientColors(params VmlGradiantColor[] colors)
        {
            if(colors==null || colors.Length==0)
            {
                throw (new ArgumentException("Please supply a list of colors"));
            }
            double p = -1;
            foreach(var c in colors)
            {
                if(c.Percent<=p)
                {
                    throw (new ArgumentException("Percent values in the color list must be sorted and must be unique."));
                }
                p = c.Percent;
            }

            var colorsString = "";
            if(colors[0].Percent!=0)
            {
                colorsString = $"0 #{colors[0].Color.ToColorString()};";
            }

            foreach(var c in colors)
            {
                var v = c.Percent == 0 ? 0 : c.Percent / 100;
                colorsString += $"{(v * 0x10000):F0}f #{c.Color.ToColorString()};";
            }
            if(colors[0].Percent==0)
            {
                _fill.Color.SetColor(colors[0].Color);
            }
            else if(!string.IsNullOrEmpty(_fill.Color.ColorString))
            {
                colorsString = $"0 #{_fill.Color.ColorString};";
            }
            
            if(colors[colors.Length-1].Percent==100)
            {
                SecondColor.SetColor(colors[colors.Length - 1].Color);
            }
            else if (!string.IsNullOrEmpty(SecondColor.ColorString))
            {
                colorsString += $"1 #{SecondColor.ColorString};";
            }
            ColorsString = colorsString;
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
