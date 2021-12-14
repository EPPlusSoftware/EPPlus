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
using System.Drawing;
using System.IO;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Information about the content, type, bounds and resolution of an image.
    /// </summary>
    public class ExcelImageInfo
    {
        internal ExcelImageInfo(byte[] image, ePictureType? type)
        {
            if (type != null && image != null)
            {
                SetImage(image, type.Value);
            }
            else
            {
                Width = 0;
                Height = 0;
                HorizontalResolution = ExcelDrawing.STANDARD_DPI;
                VerticalResolution = ExcelDrawing.STANDARD_DPI;
            }
        }
        /// <summary>
        /// The image.
        /// </summary>
        public byte[] ImageByteArray { get; private set; }
        /// <summary>
        /// The type of image.
        /// </summary>
        public ePictureType? Type 
        {
            get;
            private set;
        }
        /// <summary>
        /// The width of the image
        /// </summary>
        public double Width
        {
            get;
            private set;
        }
        /// <summary>
        /// The height of the image
        /// </summary>
        public double Height
        {
            get;
            private set;
        }
        /// <summary>
        /// The horizontal resolution of the image
        /// </summary>
        public double HorizontalResolution
        {
            get;
            private set;
        } = ExcelDrawing.STANDARD_DPI;
        /// <summary>
        /// The vertical resolution of the image
        /// </summary>
        public double VerticalResolution
        {
            get;
            private set;
        } = ExcelDrawing.STANDARD_DPI;

        public void SetImage(byte[] image, ePictureType pictureType)
        {
            ImageByteArray = image;
            Type = pictureType;
            if(pictureType==ePictureType.Wmz || 
               pictureType==ePictureType.Emz)
            {
                image = ImageReader.ExtractImage(image, out ePictureType? pt);
                if(pt.HasValue)
                {
                    throw new ArgumentException($"Image is not of type {pictureType}.", nameof(image));
                }
                else
                {
                    pictureType = pt.Value;
                }
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
                        Width = img.Width;
                        Height = img.Height;
                        HorizontalResolution = img.HorizontalResolution;
                        VerticalResolution = img.VerticalResolution;
                    }
                    catch
                    {
                        GetImageInformation(image, pictureType);
                    }                
               }
#endif
        }

        private bool GetImageInformation(byte[] image, ePictureType pictureType)
        {
            double w = 0, h = 0;
            if (ImageReader.TryGetImageBounds(pictureType, new MemoryStream(image), ref w, ref h, out double hr, out double vr))
            {
                Width = w;
                Height = h;
                HorizontalResolution = hr;
                VerticalResolution = vr;
                return true;
            }
            return false;
        }
    }
}