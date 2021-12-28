using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Vml;
using System.Drawing;
using System.IO;
namespace OfficeOpenXml
{
    public static class ExcelHeaderFooterDrawingExtension
	{

		/// <summary>
		/// Inserts a picture at the end of the text in the header or footer
		/// </summary>
		/// <param name="Picture">The image object containing the Picture</param>
		/// <param name="Alignment">Alignment. The image object will be inserted at the end of the Text.</param>
		public static ExcelVmlDrawingPicture InsertPicture(this ExcelHeaderFooterText hfText, Image Picture, PictureAlignment Alignment)
		{
			var b = ImageUtils.GetImageAsByteArray(Picture, out ePictureType type);
			using (var ms = new MemoryStream(b))
			{
				return hfText.InsertPicture(ms, type, Alignment);
			}
		}
	}
}

