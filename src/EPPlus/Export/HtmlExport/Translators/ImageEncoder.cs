/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/14/2024         EPPlus Software AB           Epplus 7.1
 *************************************************************************************************/
using OfficeOpenXml.Drawing;
using System;

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
