using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal static class ImageEncoder
    {
        static internal string EncodeImage(HtmlImage p, out ePictureType? type)
        {
            var img = p.Picture.Image;
            string encodedImage;
  
            if (img.Type == ePictureType.Emz || img.Type == ePictureType.Wmz)
            {

                encodedImage = Convert.ToBase64String(ImageReader.ExtractImage(img.ImageBytes, out type));
            }
            else
            {
                encodedImage = Convert.ToBase64String(img.ImageBytes);
                type = img.Type.Value;
            }

            return encodedImage;
        }
    }
}
