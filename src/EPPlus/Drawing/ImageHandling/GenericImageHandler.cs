
using OfficeOpenXml.Interfaces.Drawing.Image;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// The internal generic handler for image formats used in EPPlus.
    /// </summary>
    public class GenericImageHandler : IImageHandler
    {
        /// <summary>
        /// Supported types by the image handler
        /// </summary>
        public HashSet<ePictureType> SupportedTypes 
        {
            get;
        } =new HashSet<ePictureType>{ ePictureType.Jpg, ePictureType.Gif, ePictureType.Png, ePictureType.Bmp, ePictureType.Ico, ePictureType.Tif, ePictureType.Svg, ePictureType.WebP, ePictureType.Emf, ePictureType.Emz, ePictureType.Wmf, ePictureType.Wmz };

        /// <summary>
        /// The last exception that occured when calling <see cref="GetImageBounds(MemoryStream, ePictureType, out double, out double, out double, out double)"/>
        /// </summary>
        public Exception LastException { get; private set; } = null;

        /// <summary>
        /// Retreives the image bounds and resolution for an image
        /// </summary>
        /// <param name="image">The image data</param>
        /// <param name="type">Type type of image</param>
        /// <param name="width">The width of the image</param>
        /// <param name="height">The height of the image</param>
        /// <param name="horizontalResolution">The horizontal resolution in DPI</param>
        /// <param name="verticalResolution">The vertical resolution in DPI</param>
        /// <returns></returns>
        public bool GetImageBounds(MemoryStream image, ePictureType type, out double width, out double height, out double horizontalResolution, out double verticalResolution)
        {
            try
            {
                width = 0;
                height = 0;
                return ImageReader.TryGetImageBounds(type, image, ref width, ref height, out horizontalResolution, out verticalResolution);
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

        /// <summary>
        /// Returns if the handler is valid for the enviornment. 
        /// The generic image handler is valid in all environments, so it will always return true.
        /// </summary>
        /// <returns></returns>
        public bool ValidForEnvironment()
        {
            return true;
        }
    }
}
