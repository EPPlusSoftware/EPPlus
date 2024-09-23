using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
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
		public void AddPictureBmp()
		{
			var workbook = _pck.Workbook;
			var ws = workbook.Worksheets.Add("BmpImage");

			var pic = ws.Drawings.AddPicture("BmpFile", GetResourceFile("Code.bmp"));
			pic.From.Row = 0;
			pic.From.Column = 0;
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
		public void AddPictureTif()
		{
			var workbook = _pck.Workbook;
			var ws = workbook.Worksheets.Add("TifImage");

			var pic = ws.Drawings.AddPicture("TifFile", GetResourceFile("Code.tif"));
			pic.From.Row = 0;
			pic.From.Column = 0;
			pic.PreferRelativeResize = true;
			pic.LockAspectRatio = true;
		}

		[TestMethod]
		public void AddPictureFromImage()
		{
			var workbook = _pck.Workbook;
			var ws = workbook.Worksheets.Add("Image");

			var image = GetResourceFile("Vector Drawing.wmf");
			var pic = ws.Drawings.AddPicture("FromImage", image);
			pic.From.Row = 0;
			pic.From.Column = 0;
		}
		[TestMethod]
		public void AddPictureWmfFromStream()
		{
			var workbook = _pck.Workbook;
			var ws = workbook.Worksheets.Add("WmfImageStream");

			var imageStream = new FileStream(GetResourceFile("Vector Drawing.wmf").FullName, FileMode.Open, FileAccess.Read) ;
			var pic = ws.Drawings.AddPicture("wmfStream", imageStream, ePictureType.Wmf);
			pic.From.Row = 0;
			pic.From.Column = 0;
		}
		[TestMethod]
		public async Task AddPictureJpgFromStreamAsync()
		{
			var workbook = _pck.Workbook;
			var ws = workbook.Worksheets.Add("JpgImageStreamAsync");

			var imageStream = new FileStream(GetResourceFile("Test1.jpg").FullName, FileMode.Open, FileAccess.Read);
			var pic = await ws.Drawings.AddPictureAsync("jpgStreamAsync", imageStream, ePictureType.Jpg);
			pic.From.Row = 0;
			pic.From.Column = 0;
		}
		[TestMethod]
		public async Task AddPictureGifFromFileAsync()
		{
			var workbook = _pck.Workbook;
			var ws = workbook.Worksheets.Add("gifImageStreamAsync");

			var pic = await ws.Drawings.AddPictureAsync("gifStreamAsync", GetResourceFile("Test1.jpg"));
			pic.From.Row = 0;
			pic.From.Column = 0;
		}
		#region Changed Normal Font
		[TestMethod]
		public void AddNormalCalibri6()
		{
			var wb = _pck.Workbook;
			wb.Styles.NamedStyles[0].Style.Font.Size = 6;
			var ws = wb.Worksheets.Add("jpgCalibri6");
			var pic = ws.Drawings.AddPicture("jpgFile3", GetResourceFile("Test1.jpg"));
		}
		[TestMethod]
		public void AddNormalBroadway8()
		{
			var wb = _pck.Workbook;
			wb.Styles.NamedStyles[0].Style.Font.Name= "Broadway";
			wb.Styles.NamedStyles[0].Style.Font.Size = 8;
			var ws = wb.Worksheets.Add("jpgBroadway8");
			var pic = ws.Drawings.AddPicture("jpgFile3", GetResourceFile("Test1.jpg"));
		}
		[TestMethod]
		public void AddNormalBroadway16()
		{
			var wb = _pck.Workbook;
			wb.Styles.NamedStyles[0].Style.Font.Name = "Broadway";
			wb.Styles.NamedStyles[0].Style.Font.Size = 16;
			var ws = wb.Worksheets.Add("jpgBroadway16");
			var pic = ws.Drawings.AddPicture("jpgFile3", GetResourceFile("Test1.jpg"));
		}

		[TestMethod]
		public void AddNormalCalibri18()
		{
			var wb = _pck.Workbook;
			wb.Styles.NamedStyles[0].Style.Font.Size = 18;
			var ws = wb.Worksheets.Add("jpgCalibri18");
			var pic = ws.Drawings.AddPicture("jpgFile2", GetResourceFile("Test1.jpg"));			
		}
        #endregion

        [TestMethod]
        [ExpectedException(typeof(ArgumentException), "Illegal characters in path")]
        public void BrokenLinkedImageShouldFail()
        {
            using (var package = OpenPackage("LinkPic.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("emptyWS");

                var pic = sheet.Drawings.AddPicture("ImageName", "testafhkai/[/\\|stuff", PictureLocation.Link);

                SaveAndCleanup(package);
            }
        }
    }
}
