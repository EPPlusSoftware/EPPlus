using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Constants;
using OfficeOpenXml.RichData;
using OfficeOpenXml;

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
            Assert.IsNotNull(pic3, "Cell B3 picture was not present");
            // there is no picture in cell B2
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
        public void SetDrawingPic()
        {
            var path = @"C:\Users\MatsAlm\OneDrive - EPPlus Software AB\ImagesInCells\ImagesInCells2\purchase-license-thb.png";
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Drawings.AddPicture("p1", path);
        }
    }
}
