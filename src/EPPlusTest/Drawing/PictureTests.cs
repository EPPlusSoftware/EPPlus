using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Drawing
{
    [TestClass]
	public class PictureTests : TestBase
    {
		private static ExcelPackage _pck;
		[ClassInitialize]
		public static void Init(TestContext context)
		{
			InitBase();
			_pck = OpenPackage("Pictures.xlsx", true);
		}
		[ClassCleanup]
		public static void Cleanup()
		{
			SaveAndCleanup(_pck);
		}
		[TestMethod]
		public void AddPictureWmf()
		{
			var workbook = _pck.Workbook;
			var ws = workbook.Worksheets.Add("WmfImage");

			var pic = ws.Drawings.AddPicture("wmfFile", GetResourceFile("Vector Drawing.wmf"));
			pic.From.Row = 0;
			pic.From.Column = 0;
		}
		[TestMethod]
		public void AddPictureJpeg()
		{
			var workbook = _pck.Workbook;
			var ws = workbook.Worksheets.Add("jpgImage");

			var pic = ws.Drawings.AddPicture("jpgFile", GetResourceFile("Test1.jpg"));
			pic.From.Row = 0;
			pic.From.Column = 0;
		}
		[TestMethod]
		public void AddPictureGif()
		{
			var workbook = _pck.Workbook;
			var ws = workbook.Worksheets.Add("GifImage");

			var pic = ws.Drawings.AddPicture("gifFile", GetResourceFile("BitmapImage.gif"));
			pic.From.Row = 0;
			pic.From.Column = 0;
		}
		[TestMethod]
		public void AddPicturePng()
		{
			var workbook = _pck.Workbook;
			var ws = workbook.Worksheets.Add("PngImage");

			var pic = ws.Drawings.AddPicture("pngFile", GetResourceFile("EPPlus.png"));
			pic.From.Row = 0;
			pic.From.Column = 0;
		}
		[TestMethod]
		public void AddPictureEmf()
		{
			var workbook = _pck.Workbook;
			var ws = workbook.Worksheets.Add("EmfImage");

			var pic = ws.Drawings.AddPicture("emfFile", GetResourceFile("Code.emf"));
			pic.From.Row = 0;
			pic.From.Column = 0;
			pic.PreferRelativeResize = false;
			pic.LockAspectRatio = true;
		}
		[TestMethod]
		public void AddPictureFromImage()
		{
			var workbook = _pck.Workbook;
			var ws = workbook.Worksheets.Add("Image");

			var image = Image.FromFile(GetResourceFile("Vector Drawing.wmf").FullName);
			var pic = ws.Drawings.AddPicture("emfFile", image);
			pic.From.Row = 0;
			pic.From.Column = 0;
		}

	}
}
