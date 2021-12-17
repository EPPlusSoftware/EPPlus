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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Utils;
using System.Drawing;
using System.Drawing.Imaging;

namespace OfficeOpenXml.Compatibility
{
	internal class ImageCompat
	{
		internal static byte[] GetImageAsByteArray(Image image, out ePictureType type)
		{
			using (var ms = RecyclableMemory.GetStream())
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
