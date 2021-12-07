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
		}
		[ClassCleanup]
		public static void Cleanup()
		{
			SaveAndCleanup(_pck);
		}
        [TestMethod]
        public void AddJpgImageVia()
        {
            var ws = _pck.Workbook.Worksheets.Add("pictures");

            using (var ms1 = new MemoryStream(Properties.Resources.Test1JpgByteArray))
            {
                var image1 = ws.Drawings.AddPicture("jpg1", ms1, OfficeOpenXml.Drawing.ePictureType.Jpg);
            }
            using (var ms2 = new MemoryStream(Properties.Resources.VmlPatternImagePngByteArray))
            {
                var image2 = ws.Drawings.AddPicture("png1", ms2, OfficeOpenXml.Drawing.ePictureType.Png);
            }
        }
        [TestMethod]
        public void AddTestImagesToWorksheet()
        {
            var ws = _pck.Workbook.Worksheets.Add("picturesIS");

            using (var msGif = new MemoryStream(Properties.Resources.BitmapImageGif))
            {
                var imageGif = ws.Drawings.AddPicture("gif1", msGif, OfficeOpenXml.Drawing.ePictureType.Gif);
                imageGif.SetPosition(40, 0, 0, 0);
            }

            using (var msBmp = new MemoryStream(Properties.Resources.BitmapImageGif))
            {
                var imagebmp = ws.Drawings.AddPicture("bmp1", msBmp, OfficeOpenXml.Drawing.ePictureType.Bmp);
                imagebmp.SetPosition(40, 0, 10, 0);
            }

            using (var ms1 = new MemoryStream(Properties.Resources.Test1JpgByteArray))
            {
                var image1 = ws.Drawings.AddPicture("jpg1", ms1, OfficeOpenXml.Drawing.ePictureType.Jpg);
            }
            using (var ms2 = new MemoryStream(Properties.Resources.VmlPatternImagePngByteArray))
            {
                var image2 = ws.Drawings.AddPicture("png1", ms2, OfficeOpenXml.Drawing.ePictureType.Png);
                image2.SetPosition(0, 0, 10, 0);
            }
            using (var ms22 = new MemoryStream(Properties.Resources.Png2ByteArray))
            {
                var image22 = ws.Drawings.AddPicture("png2", ms22, OfficeOpenXml.Drawing.ePictureType.Png);
                image22.SetPosition(0, 0, 20, 0);
            }
            using (var ms23 = new MemoryStream(Properties.Resources.Png3ByteArray))
            {
                var image23 = ws.Drawings.AddPicture("png3", ms23, OfficeOpenXml.Drawing.ePictureType.Png);
                image23.SetPosition(0, 0, 30, 0);
            }
            using (var ms3 = new MemoryStream(Properties.Resources.CodeEmfByteArray))
            {
                var image3 = ws.Drawings.AddPicture("emf1", ms3, OfficeOpenXml.Drawing.ePictureType.Emf);
                image3.SetPosition(0, 0, 40, 0);
            }
            using (var ms4 = new MemoryStream(Properties.Resources.Svg1ByteArray))
            {
                var image4 = ws.Drawings.AddPicture("svg1", ms4, OfficeOpenXml.Drawing.ePictureType.Svg);
                image4.SetPosition(0, 0, 50, 0);
            }
            using (var ms5 = new MemoryStream(Properties.Resources.Svg2ByteArray))
            {
                var image5 = ws.Drawings.AddPicture("svg2", ms5, OfficeOpenXml.Drawing.ePictureType.Svg);
                image5.SetPosition(0, 0, 60, 0);
                image5.SetSize(25);
            }
            using (var ms6 = Properties.Resources.VectorDrawing)
            {
                var image6 = ws.Drawings.AddPicture("wmf", ms6, OfficeOpenXml.Drawing.ePictureType.Wmf);
                image6.SetPosition(0, 0, 70, 0);
            }

            using (var msTif = Properties.Resources.CodeTif)
            {
                var imageTif = ws.Drawings.AddPicture("tif1", msTif, OfficeOpenXml.Drawing.ePictureType.Tif);
                imageTif.SetPosition(0, 0, 80, 0);
            }

            using (var msTif = Properties.Resources.CodeTif)
            {
                var imageTif = ws.Drawings.AddPicture("tif1", msTif, OfficeOpenXml.Drawing.ePictureType.Tif);
                imageTif.SetPosition(0, 0, 80, 0);
            }

        }
        [TestMethod]
        public void AddWebPImages()
        {
            var ws = _pck.Workbook.Worksheets.Add("picturesWebP");

            //386*395
            using (var msWebP1 = GetImageMemoryStream("2_webp_a.webp"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("webp1", msWebP1, OfficeOpenXml.Drawing.ePictureType.WebP);
                imageWebP1.SetPosition(0, 0, 0, 0);
            }

            //386*395
            using (var msWebP1 = GetImageMemoryStream("2_webp_ll.webp"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("webp2", msWebP1, OfficeOpenXml.Drawing.ePictureType.WebP);
                imageWebP1.SetPosition(0, 0, 10, 0);
            }

            //400*400
            using (var msWebP1 = GetImageMemoryStream("animated.webp"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("webp3", msWebP1, OfficeOpenXml.Drawing.ePictureType.WebP);
                imageWebP1.SetPosition(0, 0, 20, 0);
            }

            //320*214
            using (var msWebP1 = GetImageMemoryStream("1.sm.webp"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("webp4-1", msWebP1, OfficeOpenXml.Drawing.ePictureType.WebP);
                imageWebP1.SetPosition(20, 0, 0, 0);
            }

            //320*214
            using (var msWebP1 = GetImageMemoryStream("2.sm.webp"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("webp4-2", msWebP1, OfficeOpenXml.Drawing.ePictureType.WebP);
                imageWebP1.SetPosition(20, 0, 10, 0);
            }

            //320*214
            using (var msWebP1 = GetImageMemoryStream("3.sm.webp"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("webp4-3", msWebP1, OfficeOpenXml.Drawing.ePictureType.WebP);
                imageWebP1.SetPosition(20, 0, 20, 0);
            }

            //320*214
            using (var msWebP1 = GetImageMemoryStream("4.sm.webp"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("webp4-4", msWebP1, OfficeOpenXml.Drawing.ePictureType.WebP);
                imageWebP1.SetPosition(20, 0, 30, 0);
            }

            //320*214
            using (var msWebP1 = GetImageMemoryStream("5.sm.webp"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("webp4-5", msWebP1, OfficeOpenXml.Drawing.ePictureType.WebP);
                imageWebP1.SetPosition(20, 0, 40, 0);
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
        public void AddJpgImages()
        {
            var ws = _pck.Workbook.Worksheets.Add("Jpg");

            using (var msJpg1 = GetImageMemoryStream("Jpg\\Test1.Jpg"))
            {
                var imageJpg1 = ws.Drawings.AddPicture("Jpg1", msJpg1, ePictureType.Jpg);
                imageJpg1.SetPosition(0, 0, 0, 0);
            }
        }
    }
}
