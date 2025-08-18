using EPPlusTest.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Net;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.InCellImages
{
    [TestClass]
    [DoNotParallelize]
    public class CopyCellImagesTests : TestBase
    {
        private class TestHttpsService : IHttpsService
        {
            public int NumberOfCalls { get; set; }

            public byte[] Download(string url)
            {
                NumberOfCalls++;
                return Resources.Png2ByteArray;
            }
        }

        [TestMethod]
        public void CopyCellPictureOnSameWorksheet()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            var pic = sheet.Cells["A1"].Picture.Get();
            // copy
            sheet.Cells["A1"].Copy(sheet.Cells["B1"]);
            SaveWorkbook("InCellPictureCopy_SameWorksheet.xlsx", package);
        }

        [TestMethod]
        public void CopyCellPictureToOtherWorksheet()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
            sheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            var pic = sheet.Cells["A1"].Picture.Get();
            // copy
            sheet.Cells["A1"].Copy(sheet2.Cells["B1"]);
            SaveWorkbook("InCellPictureCopy_OtherWorksheet.xlsx", package);
        }

        [TestMethod]
        public void CopyCellPictureToSheetInOtherWorkbook()
        {
            using var sourcePackage = new ExcelPackage();
            using var targetPackage = new ExcelPackage();
            var sourceSheet = sourcePackage.Workbook.Worksheets.Add("Sheet1");
            sourceSheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            var targetSheet = targetPackage.Workbook.Worksheets.Add("Sheet1");
            sourceSheet.Cells["A1"].Copy(targetSheet.Cells["A1"]);
            SaveWorkbook("CellLocalPicture_Copy_ToSheetInOtherWorkbook.xlsx", targetPackage);
        }

        [TestMethod]
        public void CopyCellPictureToSheetInOtherWorkbook_IgnoreFlag()
        {
            using var sourcePackage = new ExcelPackage();
            using var targetPackage = new ExcelPackage();
            var sourceSheet = sourcePackage.Workbook.Worksheets.Add("Sheet1");
            sourceSheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            var targetSheet = targetPackage.Workbook.Worksheets.Add("Sheet1");
            sourceSheet.Cells["A1"].Copy(targetSheet.Cells["A1"], ExcelRangeCopyOptionFlags.ExcludeLocalCellPictures);
            Assert.IsFalse(targetSheet.Cells["A1"].Picture.Exists);
        }

        [TestMethod]
        public void CopyWebPictureToSheetInOtherWorkbook()
        {
            using var sourcePackage = new ExcelPackage();
            sourcePackage.Settings.ImageFunctionService = new TestHttpsService();
            using var targetPackage = new ExcelPackage();
            var sourceSheet = sourcePackage.Workbook.Worksheets.Add("Sheet1");
            var targetSheet = targetPackage.Workbook.Worksheets.Add("Sheet1");
            sourceSheet.Cells["A1"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\")";
            sourceSheet.Calculate();
            var sourcePic = sourceSheet.Cells["A1"].Picture.Get();
            sourceSheet.Cells["A1"].Copy(targetSheet.Cells["A1"]);
            var targetPic = targetSheet.Cells["A1"].Picture.Get();
            Assert.IsNotNull(targetPic);
            Assert.AreEqual(sourcePic.FileName, targetPic.FileName);
            SaveWorkbook("CellWebPicture_Copy_ToSheetInOtherWorkbook.xlsx", targetPackage);
        }

        [TestMethod]
        public void CopyWebPictureToSheetInOtherWorkbook_IgnoreFlag()
        {
            using var sourcePackage = new ExcelPackage();
            sourcePackage.Settings.ImageFunctionService = new TestHttpsService();
            using var targetPackage = new ExcelPackage();
            var sourceSheet = sourcePackage.Workbook.Worksheets.Add("Sheet1");
            var targetSheet = targetPackage.Workbook.Worksheets.Add("Sheet1");
            sourceSheet.Cells["A1"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\")";
            sourceSheet.Calculate();
            var sourcePic = sourceSheet.Cells["A1"].Picture.Get();
            sourceSheet.Cells["A1"].Copy(targetSheet.Cells["A1"], ExcelRangeCopyOptionFlags.ExcludeWebPictures);
            var targetPic = targetSheet.Cells["A1"].Picture.Get();
            Assert.IsFalse(targetSheet.Cells["A1"].Picture.Exists);
        }


        [TestMethod]
        public void CopyCellPictureCopyEntireWorksheet()
        {
            using var package = new ExcelPackage();
            var sourceSheet = package.Workbook.Worksheets.Add("Sheet1");
            sourceSheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            package.Workbook.Worksheets.Add("Copy", sourceSheet);
            SaveWorkbook("CellLocalPicture_CopiedWorksheet.xlsx", package);
        }

        [TestMethod]
        public void CopyCellPictureCopyEntireWorksheet_ToOtherWorkbook()
        {
            using var sourcePackage = new ExcelPackage();
            var sourceSheet = sourcePackage.Workbook.Worksheets.Add("Sheet1");
            sourceSheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            using var targetPackage = new ExcelPackage();
            targetPackage.Workbook.Worksheets.Add("Copy", sourceSheet);
            SaveWorkbook("CellLocalPicture_CopiedWorksheet_OtherWorkbook.xlsx", sourcePackage);
        }
    }
}
