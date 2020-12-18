///  <v:fill color2 = "black" recolor="t" rotate="t" focus="100%" type="gradient"/>
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Xml;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils.Extensions;
namespace OfficeOpenXml.Drawing.Vml
{
    public class ExcelVmlDrawingPictureFill : XmlHelper, IPictureContainer
    {
        ExcelVmlDrawingFill _fill;
        internal ExcelVmlDrawingPictureFill(ExcelVmlDrawingFill fill, XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
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
        /// 
        /// </summary>
        public eVmlAspectRatio AspectRatio 
        { 
            get
            {
                return GetXmlNodeString("v:fill/@aspect").ToEnum(eVmlAspectRatio.Ignore);
            }
            set
            {
                SetXmlNodeString("v:fill/@aspect", value.ToString().ToLower());
            }
        }
        /// <summary>
        /// A string representing the pictures Size. 
        /// For Example: 0,0
        /// </summary>
        public string Size
        {
            get
            {
                return GetXmlNodeString("v:fill/@size");
            }
            set
            {
                SetXmlNodeString("v:fill/@size", value, true);
            }
        }
        /// <summary>
        /// A string representing the pictures Origin
        /// </summary>
        public string Origin
        {
            get
            {
                return GetXmlNodeString("v:fill/@origin");
            }
            set
            {
                SetXmlNodeString("v:fill/@origin", value, true);
            }
        }
        /// <summary>
        /// A string representing the pictures position
        /// </summary>
        public string Position
        {
            get
            {
                return GetXmlNodeString("v:fill/@position");
            }
            set
            {
                SetXmlNodeString("v:fill/@position", value, true);
            }
        }
        public string Title
        {
            get
            {
                return GetXmlNodeString("v:fill/@o:title");
            }
            set
            {
                SetXmlNodeString("v:fill/@o:title", value, true);
            }
        }
        internal Image _image = null;
        /// <summary>
        /// The image is used when <see cref="ExcelVmlDrawingFill.Style"/> is set to  Pattern, Tile or Frame.
        /// </summary>
        public Image Image
        {
            get
            {
                if(_image==null)
                {
                    var relId = GetXmlNodeString("v:fill/@o:relid");
                    if (!string.IsNullOrEmpty(relId))
                    {
                        _image = PictureStore.GetPicture(relId, this, out string contentType);
                    }
                }
                return _image;
            }
            set
            {
                if (_image == value) return;
                if (_image != null)
                {
                    var container = (IPictureContainer)this;
                    var pictureRelationDocument = (IPictureRelationDocument)_fill;
                    pictureRelationDocument.Package.PictureStore.RemoveImage(container.ImageHash, this);
                    pictureRelationDocument.RelatedPart.DeleteRelationship(container.RelPic.Id);
                    pictureRelationDocument.Hashes.Remove(container.ImageHash);
                }
                if (value != null)
                {
                    _image = value;
                }
            }
        }
        internal void SaveImage()
        {
            if (_image != null)
            {
                try
                {
                    string relId = PictureStore.SavePicture(_image, this);

                    //Create relationship
                    SetXmlNodeString("v:fill/@o:relid", relId);
                }
                catch (Exception ex)
                {
                    throw (new Exception("Can't save image - " + ex.Message, ex));
                }
            }
        }
        IPictureRelationDocument IPictureContainer.RelationDocument => _fill._drawings.Worksheet.VmlDrawings;

        string IPictureContainer.ImageHash { get; set ; }
        Uri IPictureContainer.UriPic { get; set ; }
        ZipPackageRelationship IPictureContainer.RelPic { get; set; }
    }
}
