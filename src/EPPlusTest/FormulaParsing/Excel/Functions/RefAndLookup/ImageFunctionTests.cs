using EPPlusTest.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.CellPictures;
using OfficeOpenXml.Interfaces.Net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class ImageFunctionTests : TestBase
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
        public void ImageTest_Simple()
        {
            using var package = new ExcelPackage();
            package.Settings.ImageFunctionService = new TestHttpsService();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\")";
            sheet.Calculate();
            var pic = sheet.Cells["A1"].Picture.Get();
            Assert.IsNotNull(pic, "pic was null");
            Assert.AreEqual(ExcelCellPictureTypes.WebImage, pic.PictureType);
        }

        [TestMethod]
        public void ImageTest_WithReference()
        {
            using var package = new ExcelPackage();
            package.Settings.ImageFunctionService = new TestHttpsService();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\")";
            sheet.Cells["B1"].Formula = "A1";
            sheet.Calculate();
            var pic = sheet.Cells["A1"].Picture.Get();
            Assert.IsNotNull(pic, "pic was null");
            Assert.AreEqual(ExcelCellPictureTypes.WebImage, pic.PictureType);
            Assert.AreEqual(CalcOrigins.Formula, pic.CalcOrigin);
            var pic2 = sheet.Cells["B1"].Picture.Get();
            Assert.IsNotNull(pic2);
            Assert.AreEqual(CalcOrigins.Reference, pic2.CalcOrigin);
            SaveWorkbook("ImageFunctionTest_Reference.xlsx", package);
        }
        [TestMethod]
        public void ImageTest_AltText()
        {
            using var package = new ExcelPackage();
            package.Settings.ImageFunctionService = new TestHttpsService();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\", \"Alt text\")";
            sheet.Calculate();
            var pic = sheet.Cells["A1"].Picture.Get();
            Assert.IsNotNull(pic, "pic was null");
            Assert.AreEqual(ExcelCellPictureTypes.WebImage, pic.PictureType);
            Assert.AreEqual(CalcOrigins.Formula, pic.CalcOrigin);
            Assert.AreEqual("Alt text", pic.AltText);
            SaveWorkbook("ImageFunction_AltText.xlsx", package);
        }

        [TestMethod]
        public void ImageTest_AltTextAndSizing()
        {
            using var package = new ExcelPackage();
            package.Settings.ImageFunctionService = new TestHttpsService();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\", \"Alt text\", 1)";
            sheet.Calculate();
            var pic = sheet.Cells["A1"].Picture.Get();
            Assert.IsNotNull(pic, "pic was null");
            Assert.AreEqual(ExcelCellPictureTypes.WebImage, pic.PictureType);
            Assert.AreEqual(CalcOrigins.Formula, pic.CalcOrigin);
            Assert.AreEqual("Alt text", pic.AltText);
            SaveWorkbook("ImageFunction_AltTextAndSizing.xlsx", package);
        }

        [TestMethod]
        public void ImageTest_ShouldReturnValueErrorIfHeightOrWidthIsSetAndSizingIsNotCustom()
        {
            using var package = new ExcelPackage();
            package.Settings.ImageFunctionService = new TestHttpsService();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\", \"Alt text\", 1, 1)";
            sheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ImageTest_DifferentVariants1()
        {
            using var package = new ExcelPackage();
            package.Settings.ImageFunctionService = new TestHttpsService();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\", \"Alt text\", 1)";
            sheet.Cells["B1"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\")";

            sheet.Calculate();

            var pic1 = sheet.Cells["A1"].Picture.Get();
            Assert.IsNotNull(pic1, "pic was null");
            Assert.AreEqual(ExcelCellPictureTypes.WebImage, pic1.PictureType);
            Assert.AreEqual(CalcOrigins.Formula, pic1.CalcOrigin);
            Assert.AreEqual("Alt text", pic1.AltText);

            var pic2 = sheet.Cells["B1"].Picture.Get();
            Assert.IsNotNull(pic2);
            SaveWorkbook("ImageFunction_DifferentVariants.xlsx", package);
        }

        [TestMethod]
        public void ImageTest_ShouldCacheUrls1()
        {
            using var package = new ExcelPackage();
            var httpsService = new TestHttpsService();
            package.Settings.ImageFunctionService = httpsService;
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\", \"Alt text\", 1)";
            sheet.Cells["B1"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\")";

            sheet.Calculate();

            Assert.AreEqual(1, httpsService.NumberOfCalls);
        }
    }
}
