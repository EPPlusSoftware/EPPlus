
using OfficeOpenXml.Interfaces.Image;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace OfficeOpenXml.Drawing
{
    public class GenericImageHandler : IImageHandler
    {
        public HashSet<ePictureType> SupportedTypes 
        {
            get;
        } =new HashSet<ePictureType>{ ePictureType.Jpg, ePictureType.Gif, ePictureType.Png, ePictureType.Bmp, ePictureType.Ico, ePictureType.Tif, ePictureType.Svg, ePictureType.WebP, ePictureType.Emf, ePictureType.Emz, ePictureType.Wmf, ePictureType.Wmz };

        public Exception LastException { get; private set; } = null;

        public bool GetImageBounds(MemoryStream image, ePictureType type, out double width, out double height, out double horizontalResolution, out double verticalResolution)
        {
            try
            {
                width = 0;
                height = 0;
#if(Core)
                return ImageReader.TryGetImageBounds(type, image, ref width, ref height, out horizontalResolution, out verticalResolution);
#else
            if(type==ePictureType.Ico || 
                type==ePictureType.Svg ||
                type==ePictureType.WebP)
            {
                return ImageReader.TryGetImageBounds(type, image, ref width, ref height, out horizontalResolution, out verticalResolution);
            }
            else
            {
                var img = Image.FromStream(image);
                width = img.Width;
                height = img.Height;
                horizontalResolution = img.HorizontalResolution;
                verticalResolution = img.VerticalResolution;
                return true;
            }
#endif
            }
            catch (Exception ex)
            {
                width = 0;
                height = 0;
                horizontalResolution = 0;
                verticalResolution = 0;
                LastException = ex;
                return false;
            }
        }
    }
}
