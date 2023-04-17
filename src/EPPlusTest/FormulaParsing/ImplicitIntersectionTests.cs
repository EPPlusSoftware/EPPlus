using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Packaging.Ionic.Crc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class ImplicitIntersectionTests : TestBase
    {
        [TestMethod]
        public void SingleShouldDoImplicitIntersection()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 2].Value = 1;
                sheet.Cells[2, 2].Value = 2;
                sheet.Cells[3, 2].Value = 3;
                sheet.Cells[4, 2].Value = 4;
                sheet.Cells[5, 2].Value = 5;
                sheet.Cells["A1:A5"].Formula = "B1:B5";
                sheet.Cells["C3"].Formula = "_xlfn.SINGLE(B1:B3)";
                sheet.Calculate();
                SaveAndCleanup(package);
                Assert.AreEqual(3, sheet.Cells["C3"].Value);
            }
        }
    }
}
