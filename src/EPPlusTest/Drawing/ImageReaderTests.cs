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
	public class ImageReaderTests : TestBase
    {
		private static ExcelPackage _pck;
		[ClassInitialize]
		public static void Init(TestContext context)
		{
			InitBase();
			_pck = OpenPackage("ImageReader.xlsx", true);
            _pck.Settings.ImageSettings.PrimaryImageHandler = new OfficeOpenXml.Drawing.GenericImageHandler();
        }
        [ClassCleanup]
		public static void Cleanup()
		{
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;
            
            SaveAndCleanup(_pck);
            if (File.Exists(fileName)) File.Copy(fileName, dirName + "\\ImageReaderRead.xlsx", true);
        }
        [TestMethod]
        public void AddJpgImageVia()
        {
            var ws = _pck.Workbook.Worksheets.Add("InternalJpg");

            using (var ms = new MemoryStream(Properties.Resources.Test1JpgByteArray))
            {
                var image = ws.Drawings.AddPicture("jpg", ms, OfficeOpenXml.Drawing.ePictureType.Jpg);
            }
        }
        [TestMethod]
        public void AddPngImageVia()
        {
            var ws = _pck.Workbook.Worksheets.Add("InternalPng");

            using (var ms = new MemoryStream(Properties.Resources.VmlPatternImagePngByteArray))
            {
                var image = ws.Drawings.AddPicture("png1", ms, OfficeOpenXml.Drawing.ePictureType.Png);
            }
        }

        [TestMethod]
        public void AddTestImagesToWorksheet()
        {
            var ws = _pck.Workbook.Worksheets.Add("picturesIS");

            using (var msGif = new MemoryStream(Properties.Resources.BitmapImageGif))
            {
                var imageGif = ws.Drawings.AddPicture("gif1", msGif, ePictureType.Gif);
                imageGif.SetPosition(40, 0, 0, 0);
            }

            using (var msBmp = new MemoryStream(Properties.Resources.CodeBmp))
            {
                var imagebmp = ws.Drawings.AddPicture("bmp1", msBmp, ePictureType.Bmp);
                imagebmp.SetPosition(40, 0, 10, 0);
            }

            
            using (var ms1 = new MemoryStream(Properties.Resources.Test1JpgByteArray))
            {
                var image1 = ws.Drawings.AddPicture("jpg1", ms1, ePictureType.Jpg);
            }
            
            using (var ms2 = new MemoryStream(Properties.Resources.VmlPatternImagePngByteArray))
            {
                var image2 = ws.Drawings.AddPicture("png1", ms2, ePictureType.Png);
                image2.SetPosition(0, 0, 10, 0);
            }
            
            using (var ms22 = new MemoryStream(Properties.Resources.Png2ByteArray))
            {
                var image22 = ws.Drawings.AddPicture("png2", ms22, ePictureType.Png);
                image22.SetPosition(0, 0, 20, 0);
            }
            
            using (var ms23 = new MemoryStream(Properties.Resources.Png3ByteArray))
            {
                var image23 = ws.Drawings.AddPicture("png3", ms23, ePictureType.Png);
                image23.SetPosition(0, 0, 30, 0);
            }
            
            using (var ms3 = new MemoryStream(Properties.Resources.CodeEmfByteArray))
            {
                var image3 = ws.Drawings.AddPicture("emf1", ms3, ePictureType.Emf);
                image3.SetPosition(0, 0, 40, 0);
            }

            using (var ms4 = new MemoryStream(Properties.Resources.Svg1ByteArray))
            {
                var image4 = ws.Drawings.AddPicture("svg1", ms4, ePictureType.Svg);
                image4.SetPosition(0, 0, 50, 0);
            }

            using (var ms5 = new MemoryStream(Properties.Resources.Svg2ByteArray))
            {
                var image5 = ws.Drawings.AddPicture("svg2", ms5, ePictureType.Svg);
                image5.SetPosition(0, 0, 60, 0);
                image5.SetSize(25);
            }

            using (var ms6 = Properties.Resources.VectorDrawing)
            {
                var image6 = ws.Drawings.AddPicture("wmf", ms6, ePictureType.Wmf);
                image6.SetPosition(0, 0, 70, 0);
            }

            using (var msTif = Properties.Resources.CodeTif)
            {
                var imageTif = ws.Drawings.AddPicture("tif1", msTif, ePictureType.Tif);
                imageTif.SetPosition(0, 0, 80, 0);
            }
        }
        [TestMethod]
        public void AddIcoImages()
        {
            var ws = _pck.Workbook.Worksheets.Add("Icon");
            //32*32
            using (var msIco1 = GetImageMemoryStream("1_32x32.ico"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("ico1", msIco1, OfficeOpenXml.Drawing.ePictureType.Ico);
                imageWebP1.SetPosition(40, 0, 0, 0);
            }
            //128*128
            using (var msIco2 = GetImageMemoryStream("1_128x128.ico"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("ico2", msIco2, OfficeOpenXml.Drawing.ePictureType.Ico);
                imageWebP1.SetPosition(40, 0, 10, 0);
            }

            using (var msIco3 = GetImageMemoryStream("example.ico"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("ico3", msIco3, OfficeOpenXml.Drawing.ePictureType.Ico);
                imageWebP1.SetPosition(40, 0, 20, 0);
            }

            using (var msIco4 = GetImageMemoryStream("example_small.ico"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("ico4", msIco4, OfficeOpenXml.Drawing.ePictureType.Ico);
                imageWebP1.SetPosition(40, 0, 30, 0);
            }
            using (var msIco5 = GetImageMemoryStream("Ico-file-for-testing.ico"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("ico5", msIco5, OfficeOpenXml.Drawing.ePictureType.Ico);
                imageWebP1.SetPosition(40, 0, 40, 0);
            }
        }
        [TestMethod]
        public void AddEmzImages()
        {
            var ws = _pck.Workbook.Worksheets.Add("Emz");
            //32*32
            using (var msIco1 = GetImageMemoryStream("example.emz"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("Emf", msIco1, OfficeOpenXml.Drawing.ePictureType.Emz);
                imageWebP1.SetPosition(40, 0, 0, 0);
            }
        }
        [TestMethod]
        public void AddBmpImages()
        {
            var ws = _pck.Workbook.Worksheets.Add("bmp");
            
            using (var msBmp1 = GetImageMemoryStream("bmp\\MARBLES.BMP"))
            {
                var imageBmp1 = ws.Drawings.AddPicture("bmp1", msBmp1, OfficeOpenXml.Drawing.ePictureType.Bmp);
                imageBmp1.SetPosition(0, 0, 0, 0);
            }

            using (var msBmp2 = GetImageMemoryStream("bmp\\Land.BMP"))
            {
                var imageBmp2 = ws.Drawings.AddPicture("bmp2", msBmp2, OfficeOpenXml.Drawing.ePictureType.Bmp);
                imageBmp2.SetPosition(0, 0, 20, 0);
            }

            using (var msBmp3 = GetImageMemoryStream("bmp\\Land2.BMP"))
            {
                var imageBmp3 = ws.Drawings.AddPicture("bmp3", msBmp3, OfficeOpenXml.Drawing.ePictureType.Bmp);
                imageBmp3.SetPosition(0, 0, 40, 0);
            }

            using (var msBmp4 = GetImageMemoryStream("bmp\\Land3.BMP"))
            {
                var imageBmp4 = ws.Drawings.AddPicture("bmp4", msBmp4, OfficeOpenXml.Drawing.ePictureType.Bmp);
                imageBmp4.SetPosition(0, 0, 60, 0);
            }
        }
        [TestMethod]
        public void AddWepPImages()
        {
            AddFilesToWorksheet("webp", ePictureType.WebP);
        }
        [TestMethod]
        public void AddEmfImages()
        {
            AddFilesToWorksheet("Emf", ePictureType.Emf);
        }
        [TestMethod]
        public void AddGifImages()
        {
            AddFilesToWorksheet("Gif", ePictureType.Gif);
        }
        [TestMethod]
        public void AddJpgImages()
        {
            
            AddFilesToWorksheet("Jpg", ePictureType.Jpg);
        }
        [TestMethod]
        public void AddSvgImages()
        {
            AddFilesToWorksheet("Svg", ePictureType.Svg);
        }
        [TestMethod]
        public void AddPngImages()
        {
            AddFilesToWorksheet("Png", ePictureType.Png);
        }
        [TestMethod]
        public void ReadImages()
        {
            using (var p = OpenPackage("ImageReaderRead.xlsx"))
            {
                if(p.Workbook.Worksheets.Count==0)
                {
                    Assert.Inconclusive("ImageReaderRead.xlsx does not exists. Run a full test round to create it.");
                }

                foreach(var ws in p.Workbook.Worksheets)
                {
                    ws.Columns[1, 20].Width = 35;

                    Assert.AreEqual(35, ws.Columns[1].Width);
                    Assert.AreEqual(35, ws.Columns[20].Width);
                }

                var ws2 = p.Workbook.Worksheets.Add("Bmp2");
                using (var msBmp1 = GetImageMemoryStream("bmp\\MARBLES.BMP"))
                {
                    var imageBmp1 = ws2.Drawings.AddPicture("bmp2", msBmp1, OfficeOpenXml.Drawing.ePictureType.Bmp);
                    imageBmp1.SetPosition(0, 0, 0, 0);
                }


                SaveWorkbook("ImageReaderResized.xlsx", p);
            }

        }
        [TestMethod]
        public async Task AddJpgImagesViaExcelImage()
        {
            var ws = _pck.Workbook.Worksheets.Add("AddViaExcelImage");

            var ei1 = new ExcelImage(Properties.Resources.Test1.FullName);
            Assert.IsNotNull(ei1);
            ws.BackgroundImage.Image.SetImage(ei1);

            var ei2 = new ExcelImage(Properties.Resources.Png2ByteArray, ePictureType.Png);
            Assert.IsNotNull(ei2);
            ws.BackgroundImage.Image.SetImage(ei2);

            var ei3 = new ExcelImage(new MemoryStream(Properties.Resources.BitmapImageGif), ePictureType.Gif);
            Assert.IsNotNull(ei3);

            ws.BackgroundImage.Image.SetImage(ei3);
            ws.BackgroundImage.Image.SetImage(new MemoryStream(Properties.Resources.BitmapImageGif), ePictureType.Gif);
            await ws.BackgroundImage.Image.SetImageAsync(new MemoryStream(Properties.Resources.BitmapImageGif), ePictureType.Gif);
        }

        private static void AddFilesToWorksheet(string fileType, ePictureType type)
        {
            var ws = _pck.Workbook.Worksheets.Add(fileType);

            var dir = new DirectoryInfo(_imagePath + fileType);
            if(dir.Exists==false)
            {
                Assert.Inconclusive($"Directory {dir} does not exist.");
            }
            var ix = 0;
            foreach (var f in dir.EnumerateFiles())
            {
                using (var ms = new MemoryStream(File.ReadAllBytes(f.FullName)))
                {
                    var picture = ws.Drawings.AddPicture($"{fileType}{ix}", ms, type);
                    picture.SetPosition((ix / 5) * 10, 0, (ix % 5) * 10, 0);
                    ix++;
                }
            }
         }
    }
}
