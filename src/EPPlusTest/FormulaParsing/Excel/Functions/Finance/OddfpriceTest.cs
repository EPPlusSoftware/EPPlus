using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Finance
{
    [TestClass]
    public class OddfpriceTest : TestBase
    {
        [TestMethod]
        public void OddfpriceShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with correct input values");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 02, 01);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 02, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 12, 01);
                sheet.Cells["B4"].Value = new System.DateTime(2019, 02, 15);
                //sheet.Cells["A1"].Formula = "ODDFPRICE(B1,B2,B3,B4,5%,6%,100,2,0)";
                sheet.Cells["A1"].Formula = "ODDFPRICE(B1,B2,B3,B4,1%,6%,100,2,0)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                //Assert.AreEqual(97.26007079, result);
                Assert.AreEqual(86.29690031, result);


            }
        }
    }

}