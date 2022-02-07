
#if(NETFULL)
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Only used in .NET framework.
    /// </summary>
    internal static class ImageUtils
    {
        internal static byte[] GetImageAsByteArray(Image image, out ePictureType type)
        {
            using (var ms = new MemoryStream())
            {
                if (image.RawFormat.Guid == ImageFormat.Gif.Guid)
                {
                    image.Save(ms, ImageFormat.Gif);
                    type = ePictureType.Gif;
                }
                else if (image.RawFormat.Guid == ImageFormat.Bmp.Guid)
                {
                    image.Save(ms, ImageFormat.Bmp);
                    type = ePictureType.Bmp;
                }
                else if (image.RawFormat.Guid == ImageFormat.Png.Guid)
                {
                    image.Save(ms, ImageFormat.Png);
                    type = ePictureType.Png;
                }
                else if (image.RawFormat.Guid == ImageFormat.Tiff.Guid)
                {
                    image.Save(ms, ImageFormat.Tiff);
                    type = ePictureType.Tif;
                }
                else
                {
                    image.Save(ms, ImageFormat.Jpeg);
                    type = ePictureType.Jpg;
                }

                return ms.ToArray();
            }
        }
    }
}
#endif