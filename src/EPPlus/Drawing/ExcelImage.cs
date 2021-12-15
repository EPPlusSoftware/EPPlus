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
using OfficeOpenXml.Drawing.Interfaces;
using System;
using System.IO;
#if(NETFRAMEWORK)
using System.Drawing;
#endif
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Represents an image 
    /// </summary>
    public class ExcelImage
    {
        IPictureContainer _container;
        ExcelPackage _pck;
        internal ExcelImage(IPictureContainer container)
        {
            _container = container;
            _pck = container.RelationDocument.Package;
        }

        //internal ExcelImage(byte[] image, ePictureType? type)
        //{
        //    if (type != null && image != null)
        //    {
        //        SetImage(image, type.Value);
        //    }
        //    else
        //    {
        //        Bounds.Width = 0;
        //        Bounds.Height = 0;
        //        Bounds.HorizontalResolution = ExcelDrawing.STANDARD_DPI;
        //        Bounds.VerticalResolution = ExcelDrawing.STANDARD_DPI;
        //    }
        //}
        /// <summary>
        /// The type of image.
        /// </summary>
        public ePictureType? Type
        {
            get;
            internal set;
        }

        /// <summary>
        /// The image as a byte array.
        /// </summary>
        public byte[] ImageBytes 
        { 
            get;
            internal set; 
        }
        public ExcelImageInfo Bounds
        {
            get;            
            internal set;
        } = new ExcelImageInfo();
        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="image">The image as a byte array.</param>
        /// <param name="pictureType">The type of image.</param>
        public void SetImage(byte[] image, ePictureType pictureType)
        {
            Type = pictureType;
            if (pictureType == ePictureType.Wmz ||
               pictureType == ePictureType.Emz)
            {
                var img = ImageReader.ExtractImage(image, out ePictureType? pt);
                if (pt.HasValue)
                {
                    throw new ArgumentException($"Image is not of type {pictureType}.", nameof(image));
                }
                else
                {
                    RemoveImage();
                    ImageBytes = img;
                    pictureType = pt.Value;
                }            
            }
            else
            {
                RemoveImage();
                ImageBytes = image;
            }
#if (Core)
            GetImageInformation(image, pictureType);
#else
            if(pictureType == ePictureType.Ico ||
               pictureType == ePictureType.Svg ||
               pictureType == ePictureType.WebP)
              { 
                  GetImageInformation(image, pictureType);
              }
              else
              {
                    try
                    {
                        var ms=new MemoryStream(image);
                        var img = Image.FromStream(ms);
                        Bounds.Width = img.Width;
                        Bounds.Height = img.Height;
                        Bounds.HorizontalResolution = img.HorizontalResolution;
                        Bounds.VerticalResolution = img.VerticalResolution;
                    }
                    catch
                    {
                        GetImageInformation(image, pictureType);
                    }                
               }
#endif

            _container.SetNewImage();
        }

        private void RemoveImage()
        {
            _container.RelationDocument.Package.PictureStore.RemoveImage(_container.ImageHash, _container);
            _container.RelationDocument.RelatedPart.DeleteRelationship(_container.RelPic.Id);
            _container.RelationDocument.Hashes.Remove(_container.ImageHash);
            _container.RemoveImage();
        }

        private bool GetImageInformation(byte[] image, ePictureType pictureType)
        {
            double w = 0, h = 0;
            if (ImageReader.TryGetImageBounds(pictureType, new MemoryStream(image), ref w, ref h, out double hr, out double vr))
            {
                Bounds.Width = w;
                Bounds.Height = h;
                Bounds.HorizontalResolution = hr;
                Bounds.VerticalResolution = vr;
                return true;
            }
            return false;
        }
    }
}

