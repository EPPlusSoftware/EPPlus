using EPPlusTest.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.CellPictures;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.InCellImages
{
    [TestClass]
    public class WebImagesTests : TestBase
    {
        [TestMethod]
        public void LoadSimpleWorkbook1()
        {
            using var package = OpenTemplatePackage("ImageFunction1.xlsx");
            var sheet = package.Workbook.Worksheets.First();
            var webPic = sheet.Cells["A1"].Picture.Get();
            var uri = webPic.ImageUri;

            //sheet.Cells["A2"].Blip.Set(Resources.Png2ByteArray);
            var localPic = sheet.Cells["B1"].Picture.Get();
            var lpBytes = localPic.GetImageBytes();

            var imageBytes = webPic.GetImageBytes();

            SaveWorkbook("ImageFunction1_Output.xlsx", package);
        }

        [TestMethod]
        public void WebImagesShouldBeRemovedWhenOverwritten()
        {
            using var p = new ExcelPackage();
            var sheet = p.Workbook.Worksheets.Add("Sheet");
            sheet.Cells["A1"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\")";
            sheet.Cells["A2"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\")";
            sheet.Calculate();
            Assert.IsTrue(sheet.Cells["A1"].Picture.Exists);
            Assert.IsTrue(sheet.Cells["A2"].Picture.Exists);
            Assert.AreEqual(ExcelCellPictureTypes.WebImage, sheet.Cells["A1"].Picture.Get().PictureType);
            Assert.AreEqual(ExcelCellPictureTypes.WebImage, sheet.Cells["A2"].Picture.Get().PictureType);
            Assert.AreEqual(1, p.Workbook._images.Count);
            sheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            Assert.AreEqual(2, p.Workbook._images.Count);
            sheet.Cells["A2"].Picture.Set(Resources.Png2ByteArray);
            Assert.AreEqual(1, p.Workbook._images.Count);
            SaveWorkbook("WebImages_Removed.xlsx", p);

        }
    }
}
