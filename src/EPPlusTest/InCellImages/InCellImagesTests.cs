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

namespace EPPlusTest.InCellImages
{
    [TestClass]
    public class InCellImagesTests : TestBase
    {
        [TestMethod]
        public void GetCellPicture()
        {
            using var package = OpenTemplatePackage("InCellImage1.xlsx");

            var pic1 = package.Workbook.Worksheets[0].Cells["A1"].GetCellPicture();
            var pic2 = package.Workbook.Worksheets[0].Cells["A2"].GetCellPicture();
            var pic3 = package.Workbook.Worksheets[0].Cells["B1"].GetCellPicture();
            var pic4 = package.Workbook.Worksheets[0].Cells["B2"].GetCellPicture();

            Assert.IsNotNull(pic1, "Cell A1 picture was not present");
            Assert.IsNotNull(pic2, "Cell A2 picture was not present");
            Assert.IsNotNull(pic3, "Cell B3 picture was not present");            // there is no picture in cell B2
            Assert.IsNull(pic4, "Cell B2 was not empty");
        }

        [TestMethod]
        public void SetCellPicture()
        {
            var package = OpenPackage("InCellPictureSetPng.xlsx", delete: true);
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            var imageBytes = Resources.Png2ByteArray;
            sheet.Cells["A1"].SetCellPicture(imageBytes);
            SaveWorkbook("InCellPictureSetPng.xlsx", package);
        }

        [TestMethod]
        public void OverwriteCellPicture()
        {
            var pic1Bytes = Resources.Png2ByteArray;
            var pic2Bytes = Resources.Png3ByteArray;
            using var package = OpenPackage("InCellPictureOverwrite.xlsx", delete: true);
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].SetCellPicture(pic1Bytes);
            var pic1 = package.Workbook.Worksheets[0].Cells["A1"].GetCellPicture();
            Assert.IsNotNull(pic1, "Cell A1 picture was not present");
            sheet.Cells["A1"].SetCellPicture(pic2Bytes);
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
            sheet.Cells["A1"].SetCellPicture(pictureBytes, "This is an alt-text");
            SaveWorkbook("InCellPicturesAlt1.xlsx", package);
        }

        [TestMethod]
        public void SetCellPictureMarkAsDecorative()
        {
            using var package = OpenPackage("InCellPicturesDecorative.xlsx", delete: true);
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            var pictureBytes = Resources.CodeBmp;
            sheet.Cells["A1"].SetCellPicture(pictureBytes, markAsDecorative: true);
            SaveWorkbook("InCellPicturesDecorative.xlsx", package);
        }

        [TestMethod]
        public void PreserveGeoDataType()
        {
            using var package = OpenTemplatePackage("RichDataPreserve1.xlsx");
            var path = @"c:\Temp\RichDataPreserve1.xlsx";
            if(File.Exists(path))
            {
                File.Delete(path);
            }
            package.SaveAs(path);
        }

        [TestMethod, Ignore]
        public void TestImageFormats()
        {
            var imageDirectory = @"C:\Users\MatsAlm\dev\EPPlusSoftware\Pics";
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");

            //var images = new List<string> { "jpg1.jpg", "png1.png", "gif1.gif", "bmp1.bmp", "ico1.ico", "tif1.tif", "emf1.emf", "wmf1.wmf" };
            // doesn't work: emf, wmf, svg
            var images = new List<string> { "jpg1.jpg", "png1.png", "gif1.gif", "bmp1.bmp", "ico1.ico", "tif1.tif", "webp1.webp"};
            //var images = new List<string> { "svg1.svg" };
            for (var i = 1; i <= images.Count; i++)
            {
                sheet.Cells[i, 1].Value = images[i - 1];
                sheet.Cells[i, 2].SetCellPicture(Path.Combine(imageDirectory, images[i - 1]));
            }
            package.SaveAs(@"c:\temp\CellPictureEPPlusImageTypes.xlsx");
        }
    }
}
