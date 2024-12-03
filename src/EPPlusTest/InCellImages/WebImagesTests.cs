using EPPlusTest.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
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

            //sheet.Cells["A2"].Picture.Set(Resources.Png2ByteArray);
            var localPic = sheet.Cells["B1"].Picture.Get();
            var lpBytes = localPic.GetImageBytes();

            var imageBytes = webPic.GetImageBytes();

            SaveWorkbook("ImageFunction1_Output.xlsx", package);
        }
    }
}
