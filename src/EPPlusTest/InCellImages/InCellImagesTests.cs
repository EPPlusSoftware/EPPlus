using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Constants;
using OfficeOpenXml.RichData;
using OfficeOpenXml;
using System.IO;
using EPPlusTest.Properties;
using OfficeOpenXml.CellPictures;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System.Security.Cryptography;

namespace EPPlusTest.InCellImages
{
    [TestClass]
    [DoNotParallelize]
    public class InCellImagesTests : TestBase
    {
        [TestMethod]
        public void GetCellPicture()
        {
            using var package = OpenTemplatePackage("InCellImage1.xlsx");

            var pic1 = package.Workbook.Worksheets[0].Cells["A1"].Picture.Get();
            var pic2 = package.Workbook.Worksheets[0].Cells["A2"].Picture.Get();
            var pic3 = package.Workbook.Worksheets[0].Cells["B1"].Picture.Get();
            var pic4 = package.Workbook.Worksheets[0].Cells["B2"].Picture.Get();

            Assert.IsNotNull(pic1, "Cell A1 picture was not present");
            Assert.IsNotNull(pic2, "Cell A2 picture was not present");
            Assert.IsNotNull(pic3, "Cell B3 picture was not present");            // there is no picture in cell B2
            Assert.IsNull(pic4, "Cell B2 was not empty");

            var name1 = pic1.FileName;
            Assert.AreEqual("image1.png", name1);
            var bytes1 = pic1.GetImageBytes();
            Assert.AreEqual(12185, bytes1.Length);

            var name2 = pic2.FileName;
            Assert.AreEqual("image2.png", name2);
            var bytes2 = pic2.GetImageBytes();
            Assert.AreEqual(11306, bytes2.Length);

        }

        [TestMethod]
        public void SetCellPicture()
        {
            var package = OpenPackage("InCellPictureSetPng.xlsx", delete: true);
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            var imageBytes = Resources.Png2ByteArray;
            sheet.Cells["A1"].Picture.Set(imageBytes);
            var val = sheet.Cells["A1"].Value;
            SaveWorkbook("InCellPictureSetPng.xlsx", package);
        }

        [TestMethod]
        public void OverwriteCellPicture()
        {
            var pic1Bytes = Resources.Png2ByteArray;
            var pic2Bytes = Resources.Png3ByteArray;
            using var package = OpenPackage("InCellPictureOverwrite.xlsx", delete: true);
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Picture.Set(pic1Bytes);
            var pic1 = package.Workbook.Worksheets[0].Cells["A1"].Picture.Get();
            Assert.IsNotNull(pic1, "Cell A1 picture was not present");
            sheet.Cells["A1"].Picture.Set(pic2Bytes);
            sheet.Row(1).Height = 25;
            sheet.Column(1).Width = 50;
            SaveWorkbook("InCellPictureOverwrite.xlsx", package);
        }

        [TestMethod]
        public void SetCellPictureWithAltText()
        {
            using var package = OpenPackage("InCellPicturesAlt1.xlsx", delete: true);
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            var pictureBytes = Resources.Test1JpgByteArray;
            sheet.Cells["A1"].Picture.Set(pictureBytes, "This is an alt-text");
            SaveWorkbook("InCellPicturesAlt1.xlsx", package);
        }

        [TestMethod]
        public void SetCellPictureMarkAsDecorative()
        {
            using var package = OpenPackage("InCellPicturesDecorative.xlsx", delete: true);
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            var pictureBytes = Resources.CodeBmp;
            sheet.Cells["A1"].Picture.Set(pictureBytes, isDecorative: true);
            SaveWorkbook("InCellPicturesDecorative.xlsx", package);
        }

        [TestMethod]
        public void PreserveWithOtherGeoDataType()
        {
            using var package = OpenTemplatePackage("RichDataPreserve1.xlsx");
            var sheet = package.Workbook.Worksheets.First();
            sheet.Cells["F1"].Picture.Set(Resources.Png2ByteArray);
            SaveWorkbook("InCellImageWithOtherRichDataPreserve1.xlsx", package);
        }

        [TestMethod]
        public void PreserveWithOtherGeoDataTypeDeleteImage()
        {
            using var package = OpenTemplatePackage("RichDataPreserve1.xlsx");
            var sheet = package.Workbook.Worksheets.First();
            sheet.Cells["F2"].Picture.Set(Resources.Png3ByteArray);
            sheet.Cells["F1"].Picture.Set(Resources.Png2ByteArray);
            sheet.Cells["F1"].Picture.Remove();
            SaveWorkbook("InCellImageWithOtherRichDataPreserve2.xlsx", package);
        }

        [TestMethod]
        public void AddToNewPackage()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            var pic1 = sheet.Cells["A1"].Picture.Get();
            Assert.AreEqual("image1.png", pic1.FileName);
            Assert.AreEqual(Resources.Png2ByteArray.Length, pic1.GetImageBytes().Length);
        }

        [TestMethod, Ignore]
        public void TestImageFormats()
        {
            var imageDirectory = @"C:\Users\MatsAlm\dev\EPPlusSoftware\Pics";
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");

            //var images = new List<string> { "jpg1.jpg", "png1.png", "gif1.gif", "bmp1.bmp", "ico1.ico", "tif1.tif", "emf1.emf", "wmf1.wmf" };
            // doesn't work: emf, wmf, svg
            var images = new List<string> { "jpg1.jpg", "png1.png", "gif1.gif", "bmp1.bmp", "ico1.ico", "tif1.tif", "webp1.webp" };
            //var images = new List<string> { "svg1.svg" };
            for (var i = 1; i <= images.Count; i++)
            {
                sheet.Cells[i, 1].Value = images[i - 1];
                sheet.Cells[i, 2].Picture.Set(Path.Combine(imageDirectory, images[i - 1]));
            }
            package.SaveAs(@"c:\temp\CellPictureEPPlusImageTypes.xlsx");

            var font = new MeasurementFont
            {
                FontFamily = "Times New Roman",
                Size = 7,
                Style = MeasurementFontStyles.Bold
            };
            var measurement = package.Settings.TextSettings.PrimaryTextMeasurer.MeasureText("Lorem ipsum\ndolor sit amet", font);
            var widthInPixels = measurement.Width;
            var heightInPixels = measurement.Height;
        }



        [TestMethod]
        public void testImageInCell()
        {
            using var p = OpenTemplatePackage("MyRichData.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var stockholm = ws.Cells["G1"].Picture.Get();
            var text = ws.Cells["F1"].Value;
            var text2 = ws.Cells["H1"].Value;
            var tokyo = ws.Cells["G2"].Picture.Get();
            var dallas = ws.Cells["G3"].Picture.Get();

            var ws2 = p.Workbook.Worksheets.Add("Sheet 2");
            ws2.Cells["G1"].Picture.Set(stockholm.GetImageBytes());
            ws2.Cells["G2"].Picture.Set(tokyo.GetImageBytes());
            ws2.Cells["G3"].Picture.Set(dallas.GetImageBytes());

            SaveAndCleanup(p);
        }

        [TestMethod]
        public void VerifyLoadOfImageFunction()
        {
            using var p = OpenTemplatePackage("5.7-InCellPictures.xlsx");
            var ws = p.Workbook.Worksheets[0];

            ws.Calculate(x => x.AlwaysRefreshImageFunction = true);

            Assert.IsTrue(ws.Cells["A1"].Picture.Exists);
            Assert.IsTrue(ws.Cells["B1"].Picture.Exists);
            Assert.IsTrue(ws.Cells["B2"].Picture.Exists);

            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CellPictureIssue1()
        {
            using var p = OpenTemplatePackage("CpIssue1.xlsx");
            var sheet = p.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Picture.Set(Resources.Png3ByteArray);
            sheet.Cells["B1"].Formula = "A1";
            sheet.Cells["C1"].SetFormula("Image(\"https://samples.epplussoftware.com/img/EPPlus-logo-full.png\")");
            sheet.Calculate();

            sheet.Cells["A1:B1"].Picture.Remove();
            SaveAndCleanup(p);
        }

        [TestMethod]
        public void InCellPicturesMultiCellRange()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");

            var myPic = Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp");
            var imageBytes = File.ReadAllBytes(myPic);


            ws.Cells["A1:D20"].Picture.Set(imageBytes);
            SaveWorkbook("IncellPictureTestFor8.xlsx", p);
        }

        [TestMethod]
        public void PictureSet()
        {
            using var p = new ExcelPackage();
            var sheet = p.Workbook.Worksheets.Add("Sheet");
            sheet.Cells["A1"].Picture.Set(Resources.Png3ByteArray);
        }


        [TestMethod]
        public void InCellPicture_CopyBecomesLeftAlign()
        {
            using (ExcelPackage package = OpenPackage("PicturesInCellRef.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add("NewSheet");
                var wsOther = wb.Worksheets.Add("NewSheet2");
                var fi = GetResourceFile("EPPlus.png");
                ws.Cells["D4"].Picture.Set(fi);
                ws.Cells["A1:M30"].Copy(wsOther.Cells["A1:M30"]);
                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void ChangeImages()
        {
            using (ExcelPackage package = OpenPackage("CellPictures.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add("cellSheet");
                var ws2 = wb.Worksheets.Add("Ws_other");

                var fi = GetResourceFile("EPPlus.png");
                var fi2 = GetResourceFile("Test1.jpg");

                ws.Cells["F5"].Picture.Set(fi);
                ws.Cells["F6"].Picture.Set(fi);

                var pic = ws.Cells["F5"].Picture;
                pic.Set(fi2);

                var picAlt = ws.Cells["F6"].Picture.Get();
                ws.Cells["F6"].Picture.Set(fi2);

                //Maybe not quite the right assert
                //Ensure media folder only contains one file after save.
                Assert.AreEqual(1, wb._images.Count());

                SaveAndCleanup(package);
            }
        }
    }
}
