using OfficeOpenXml.Drawing;
using OfficeOpenXml.Interfaces.Image;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
namespace OfficeOpenXml.System.Drawing.Image
{
    internal class SystemDrawingImageHandler : IImageHandler
    {
        public HashSet<ePictureType> SupportedTypes
        {
            get;
        } = new HashSet<ePictureType>() { ePictureType.Bmp, ePictureType.Jpg, ePictureType.Gif, ePictureType.Png, ePictureType.Tif, ePictureType.Emf, ePictureType.Wmf };

        public Exception LastException { get; private set; }

        public bool GetImageBounds(MemoryStream image, ePictureType type, out double width, out double height, out double horizontalResolution, out double verticalResolution)
        {
            try
            {
                var bmp = new Bitmap(image);
                width = bmp.Width;
                height = bmp.Height;
                horizontalResolution = bmp.HorizontalResolution;
                verticalResolution = bmp.VerticalResolution;
                return true;
            }
            catch(Exception ex)
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
