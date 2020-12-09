///  <v:fill color2 = "black" recolor="t" rotate="t" focus="100%" type="gradient"/>
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Xml;
using OfficeOpenXml.Utils.Extensions;
namespace OfficeOpenXml.Drawing.Vml
{
    public class ExcelVmlDrawingBlipFill : XmlHelper
    {
        ExcelVmlDrawingFill _fill;
        internal ExcelVmlDrawingBlipFill(ExcelVmlDrawingFill fill, XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
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
        /// <summary>
        /// The image is used when <see cref="ExcelVmlDrawingFill.Style"/> is set to  Pattern, Tile or Frame.
        /// </summary>
        public Image Image
        {
            get;
            private set;
        }
        byte[] _image;
        public void SetImage(Image image)
        {
            Image = image;
            using (var ms = new MemoryStream())
            {
                image.Save(ms, image.RawFormat);
                _image = ms.ToArray();
            }
        }
        public void SetImage(byte[] image)
        {
            _image = image;
        }
    }
}
