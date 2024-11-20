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
            var pic = sheet.Cells["A1"].Picture.Get();
            var uri = pic.ImageUri;
        }
    }
}
