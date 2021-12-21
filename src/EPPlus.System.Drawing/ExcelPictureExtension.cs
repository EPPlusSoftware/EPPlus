using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace OfficeOpenXml.Drawing
{
    public static class ExcelPictureExtension
    {
        /// <summary>
		/// Sets the image using a System.Drawing.Image
		/// </summary>
		/// <param name="value"></param>
		/// <param name="image">The image</param>
		public static void SetImage(this ExcelImage value, Image image)
        {
			var b=GetImageAsByteArray(image, out ePictureType type);
			value.SetImage(b, type);
        }
		/// <summary>
		/// Adds a picture to the worksheet
		/// </summary>
		/// <param name="Name">The name of the drawing object</param>
		/// <param name="image">An image.</param>
		/// <returns></returns>
		public static ExcelPicture AddPicture(this ExcelDrawings drawings, string Name, Image Image)
		{
			if (Image != null)
			{
				var b = GetImageAsByteArray(Image, out ePictureType type);
				return drawings.AddPicture(Name, new MemoryStream(b), type, null);
			}
			throw (new Exception("AddPicture: Image can't be null"));
		}
		/// <summary>
		/// Adds a picture to the worksheet
		/// </summary>
		/// <param name="Name">The name of the drawing object</param>
		/// <param name="Image">An image. </param>
		/// <param name="Hyperlink">Picture Hyperlink</param>
		/// <returns>A picture object</returns>
		public static ExcelPicture AddPicture(this ExcelDrawings drawings, string Name, Image Image, Uri Hyperlink)
		{
			if (Image != null)
			{
				var b = GetImageAsByteArray(Image, out ePictureType type);
				return drawings.AddPicture(Name, new MemoryStream(b), type, Hyperlink);
			}
			throw (new Exception("AddPicture: Image can't be null"));
		}
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

