﻿using System;
using System.Xml;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils.Extensions;
namespace OfficeOpenXml.Drawing.Vml
{
    /// <summary>
    /// Fill settings for a vml pattern or picture fill
    /// </summary>
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
        /// The aspect ratio 
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
        /// <summary>
        /// The title for the fill
        /// </summary>
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
        ExcelImage _image=null;
        /// <summary>
        /// The image is used when <see cref="ExcelVmlDrawingFill.Style"/> is set to  Pattern, Tile or Frame.
        /// </summary>
        public ExcelImage Image
        {
            get
            {
                if (_image == null)
                {
                    var relId = RelId;
                    _image = new ExcelImage(this, new ePictureType[] { ePictureType.Svg, ePictureType.Ico, ePictureType.WebP });
                    if (!string.IsNullOrEmpty(relId))
                    {
                        _image.ImageBytes = PictureStore.GetPicture(relId, this, out string contentType, out ePictureType pictureType);
                        _image.Type = pictureType;
                    }
                }
                return _image;
            }
        }

        IPictureRelationDocument IPictureContainer.RelationDocument => _fill._drawings.Worksheet.VmlDrawings;

        string IPictureContainer.ImageHash { get; set ; }
        Uri IPictureContainer.UriPic { get; set ; }
        ZipPackageRelationship IPictureContainer.RelPic { get; set; }
        void IPictureContainer.SetNewImage()
        {
            var container = (IPictureContainer)this;
            //Create relationship
            SetXmlNodeString("v:fill/@o:relid", container.RelPic.Id);
        }
        void IPictureContainer.RemoveImage()
        {
            var container = (IPictureContainer)this;
            var pictureRelationDocument = (IPictureRelationDocument)_fill._drawings;
            pictureRelationDocument.Package.PictureStore.RemoveImage(container.ImageHash, this);
            pictureRelationDocument.RelatedPart.DeleteRelationship(container.RelPic.Id);
            pictureRelationDocument.Hashes.Remove(container.ImageHash);
        }

        internal string RelId 
        { 
            get
            {
                return GetXmlNodeString("v:fill/@o:relid");
            }
        }
    }
}
