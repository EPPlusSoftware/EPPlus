/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Xml;
using System.Drawing;
using System.IO;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Drawing.Interfaces;

namespace OfficeOpenXml
{
    /// <summary>
    /// An image that fills the background of the worksheet.
    /// </summary>
    public class ExcelBackgroundImage : XmlHelper, IPictureContainer
    {
        ExcelWorksheet _workSheet;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="nsm"></param>
        /// <param name="topNode">The topnode of the worksheet</param>
        /// <param name="workSheet">Worksheet reference</param>
        internal  ExcelBackgroundImage(XmlNamespaceManager nsm, XmlNode topNode, ExcelWorksheet workSheet) :
            base(nsm, topNode)
        {
            _workSheet = workSheet;
        }
        ExcelImage _imageNew;
        const string BACKGROUNDPIC_PATH = "d:picture/@r:id";
        /// <summary>
        /// The background image of the worksheet. 
        /// Note that images of type .svg, .ico and .webp is not supported as background images.
        /// </summary>
        public ExcelImage Image 
        { 
            get
            {
                if (_imageNew == null)
                {
                    var relId = GetXmlNodeString(BACKGROUNDPIC_PATH);
                    _imageNew = new ExcelImage(this, new ePictureType[] {ePictureType.Svg, ePictureType.Ico, ePictureType.WebP});
                    if (!string.IsNullOrEmpty(relId))
                    {
                        _imageNew.ImageBytes = PictureStore.GetPicture(relId, this, out string contentType, out ePictureType pictureType);
                        _imageNew.Type = pictureType;
                    }
                }
                return _imageNew;
            }
        }
        /// <summary>
        /// Set the picture from an image file. 
        /// </summary>
        /// <param name="PictureFile">The image file. Files of type .svg, .ico and .webp is not supported for background images</param>
        public void SetFromFile(FileInfo PictureFile)
        {
            if(PictureFile.Exists==false)
            {
                throw new FileNotFoundException($"Can't find file {PictureFile.FullName}");
            }
            var type = PictureStore.GetPictureType(PictureFile.Extension);
            var imgBytes =File.ReadAllBytes(PictureFile.FullName);
            Image.SetImage(imgBytes, type);
        }
        /// <summary>
        /// Set the picture from an image file. 
        /// </summary>
        /// <param name="PictureFilePath">The path to the image file. Files of type .svg, .ico and .webp is not supported for background images</param>
        public void SetFromFile(string PictureFilePath)
        {
            if (string.IsNullOrEmpty(PictureFilePath))
            {
                throw new ArgumentNullException("File path cannot be null.");
            }
            SetFromFile(new FileInfo(PictureFilePath));
        }

        IPictureRelationDocument IPictureContainer.RelationDocument 
        { 
            get 
            { 
                return _workSheet; 
            } 
        }
        string IPictureContainer.ImageHash { get; set; }
        Uri IPictureContainer.UriPic { get; set; }
        Packaging.ZipPackageRelationship IPictureContainer.RelPic { get; set; }


        void IPictureContainer.RemoveImage()
        {
            if (Image.Type != null)
            {
                var pc = (IPictureContainer)this;
                _workSheet._package.PictureStore.RemoveImage(pc.ImageHash, pc);
                _workSheet.DeleteNode(BACKGROUNDPIC_PATH, true);
            }
        }

        void IPictureContainer.SetNewImage()
        {
            var pc = (IPictureContainer)this;
            _workSheet.SetXmlNodeString(BACKGROUNDPIC_PATH, pc.RelPic.Id);
        }
    }
}



