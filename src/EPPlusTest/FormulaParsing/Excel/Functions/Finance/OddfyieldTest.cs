using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Finance
{
    [TestClass]
    public class OddfyieldTest : TestBase
    {
        [TestMethod]
        public void OddfyieldShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with correct input values");
                sheet.Cells["B1"].Value = new System.DateTime(2008, 11, 11);
                sheet.Cells["B2"].Value = new System.DateTime(2021, 03, 01);
                sheet.Cells["B3"].Value = new System.DateTime(2008, 10, 15);
                sheet.Cells["B4"].Value = new System.DateTime(2009, 03, 01);
                sheet.Cells["A1"].Formula = "ODDFPRICE(B1,B2,B3,B4,7.85%,6.25%,100,2,1)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 2);
                Assert.AreEqual(113.60, result);
            }
        }
    }

}
