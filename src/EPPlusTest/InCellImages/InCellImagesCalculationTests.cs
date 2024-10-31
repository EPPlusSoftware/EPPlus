using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using EPPlusTest.Properties;
using OfficeOpenXml.CellPictures;

namespace EPPlusTest.InCellImages
{
    [TestClass]
    public class InCellImagesCalculationTests : TestBase
    {
        [TestMethod]
        public void ShouldAddRichDataWhenReferenceFromOtherCell()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet 1");
            sheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            sheet.Cells["B1"].Formula = "A1";
            sheet.Calculate();
            var rd = package.Workbook.RichData;
            Assert.AreEqual(2, rd.Values.Count, "RichData values count was not 2");
            var cellVal = sheet.Cells["B1"].Value;
            Assert.IsNotNull(cellVal, "B1 cell value was null");
            Assert.IsInstanceOfType(cellVal, typeof(ExcelCellPicture), "Cell B1 did not contain an instance of ExcelCellPicture");
            var b1Pic = cellVal as ExcelCellPicture;
            Assert.AreEqual(CalcOrigins.Reference, b1Pic.CalcOrigin, "CalcOrigin of the B1 picture was not Reference");
            var b1Bytes = b1Pic.GetImageBytes();
            Assert.AreEqual(Resources.Png2ByteArray.Length, b1Bytes.Length, "Image bytes of the cell picture was not the same as the added picture");
            Assert.AreEqual(1, package.Workbook.RichData.RichValueRels.Count);
            SaveWorkbook("InCellImageCalculate1.xlsx", package);
        }

        [TestMethod]
        public void ShouldCauseValueErrorInSum()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet 1");
            sheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            sheet.Cells["A2"].Value = 1;
            sheet.Cells["A3"].Formula = "SUM(A1:A2)";
            sheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["A3"].Value);
        }

        [TestMethod]
        public void ShouldAddPictureInIfFunction()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet 1");
            sheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            sheet.Cells["A2"].Value = 1;
            sheet.Cells["A3"].Formula = "IF(TRUE(),A1,A2)";
            sheet.Calculate();
            Assert.IsTrue(sheet.Cells["A3"].Picture.Exists, "No picture present in cell A3 after calc");
            var pic = sheet.Cells["A3"].Picture.Get();
            Assert.AreEqual(Resources.Png2ByteArray.Length, pic.GetImageBytes().Length, "Length of A3 image bytes was not the same as the original picture");
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["A3"].Value);
            SaveWorkbook("InCellImageCalculate2.xlsx", package);
        }

        [TestMethod]
        public void ShouldRemoveCalculatedImage1()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet 1");
            sheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            sheet.Cells["A2"].Value = 1;
            sheet.Cells["A3"].Formula = "IF(TRUE(),A1,A2)";
            sheet.Calculate();
            Assert.IsTrue(sheet.Cells["A3"].Picture.Exists, "No picture present in cell A3 after calc");
            var pic = sheet.Cells["A3"].Picture.Get();
            Assert.AreEqual(Resources.Png2ByteArray.Length, pic.GetImageBytes().Length, "Length of A3 image bytes was not the same as the original picture");
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["A3"].Value);
            Assert.AreEqual(2, package.Workbook.RichData.Values.Count);
            sheet.Cells["A3"].Formula = "IF(FALSE(),A1,A2)";
            sheet.Calculate();
            Assert.AreEqual(1, sheet.Cells["A3"].Value, "Value of A3 was not 1 as expected");
            Assert.AreEqual(1, package.Workbook.RichData.Values.Count, "RichDatat.Values.Count was not 1 as expected");
            Assert.IsFalse(sheet.Cells["A3"].Picture.Exists);
            SaveWorkbook("InCellImageCalculate3.xlsx", package);
        }
    }
}
