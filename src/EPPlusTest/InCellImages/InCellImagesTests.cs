using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Constants;
using OfficeOpenXml.RichData;
using OfficeOpenXml;
using System.IO;

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
            var path = @"C:\Users\MatsAlm\OneDrive - EPPlus Software AB\ImagesInCells\ImagesInCells2\purchase-license-thb.png";
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].SetCellPicture(path);
            var pic1 = package.Workbook.Worksheets[0].Cells["A1"].GetCellPicture();
            Assert.IsNotNull(pic1, "Cell A1 picture was not present");
            package.SaveAs(@"c:\temp\CellPictureEPPlus.xlsx");
        }

        [TestMethod]
        public void OverwriteCellPicture()
        {
            var path = @"C:\Users\MatsAlm\OneDrive - EPPlus Software AB\ImagesInCells\ImagesInCells2\purchase-license-thb.png";
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].SetCellPicture(path);
            var pic1 = package.Workbook.Worksheets[0].Cells["A1"].GetCellPicture();
            Assert.IsNotNull(pic1, "Cell A1 picture was not present");
            sheet.Cells["A1"].SetCellPicture(path);
            package.SaveAs(@"c:\temp\CellPictureEPPlusOverwrite.xlsx");
        }

        [TestMethod]
        public void SetCellPictureWithAltText()
        {
            var path = @"C:\Users\MatsAlm\OneDrive - EPPlus Software AB\ImagesInCells\ImagesInCells2\purchase-license-thb.png";
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].SetCellPicture(path, "This is an alt-text");
            package.SaveAs(@"c:\temp\CellPictureEPPlusAlt1.xlsx");
        }

        [TestMethod]
        public void TestImageFormats()
        {
            var imageDirectory = @"C:\Users\MatsAlm\dev\EPPlusSoftware\Pics";
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");

            //var images = new List<string> { "jpg1.jpg", "png1.png", "gif1.gif", "bmp1.bmp", "ico1.ico", "tif1.tif", "emf1.emf", "wmf1.wmf" };
            // doesn't work: emf, wmf
            //var images = new List<string> { "jpg1.jpg", "png1.png", "gif1.gif", "bmp1.bmp", "ico1.ico", "tif1.tif", "webp1.webp"};
            var images = new List<string> { "svg1.svg" };
            for (var i = 1; i <= images.Count; i++)
            {
                sheet.Cells[i, 1].Value = images[i - 1];
                sheet.Cells[i, 2].SetCellPicture(Path.Combine(imageDirectory, images[i - 1]));
            }
            package.SaveAs(@"c:\temp\CellPictureEPPlusImageTypes.xlsx");
        }
    }
}
