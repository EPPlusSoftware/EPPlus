using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Vml;
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
			var b=ImageUtils.GetImageAsByteArray(image, out ePictureType type);
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
				var b = ImageUtils.GetImageAsByteArray(Image, out ePictureType type);
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
				var b = ImageUtils.GetImageAsByteArray(Image, out ePictureType type);
				return drawings.AddPicture(Name, new MemoryStream(b), type, Hyperlink);
			}
			throw (new Exception("AddPicture: Image can't be null"));
		}
    }
}

