using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Finance
{
    [TestClass]
    public class PriceDiscTest : TestBase
    {
        [TestMethod]
        public void PriceDiscShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with correct input values");
                sheet.Cells["B1"].Value = new System.DateTime(2008, 2, 16);
                sheet.Cells["B2"].Value = new System.DateTime(2008, 3, 1);
                sheet.Cells["A1"].Formula = "PRICEDISC(B1,B2,5.25%,100,2)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                Assert.AreEqual(99.79583333d, result);

            }
        }

        [TestMethod]
        public void PriceDiscIncorrectBasis()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect basis");
                sheet.Cells["B1"].Value = new System.DateTime(2008, 2, 16);
                sheet.Cells["B2"].Value = new System.DateTime(2008, 3, 1);
                sheet.Cells["A1"].Formula = "PRICEDISC(B1,B2,5.25%,100,-2)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);

            }
        }

        [TestMethod]
        public void PriceDiscIncorrectDiscount()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect discount");
                sheet.Cells["B1"].Value = new System.DateTime(2008, 2, 16);
                sheet.Cells["B2"].Value = new System.DateTime(2008, 3, 1);
                sheet.Cells["A1"].Formula = "PRICEDISC(B1,B2,-5.25%,100,2)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);

            }
        }

        [TestMethod]
        public void PriceDiscIncorrectRedemption()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect redemption");
                sheet.Cells["B1"].Value = new System.DateTime(2008, 2, 16);
                sheet.Cells["B2"].Value = new System.DateTime(2008, 3, 1);
                sheet.Cells["A1"].Formula = "PRICEDISC(B1,B2,5.25%,-100,2)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);

            }
        }

        [TestMethod]
        public void PriceDiscSettlementGreaterThanMaturity()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with settlement > maturity");
                sheet.Cells["B1"].Value = new System.DateTime(2008, 2, 16);
                sheet.Cells["B2"].Value = new System.DateTime(2008, 1, 1);
                sheet.Cells["A1"].Formula = "PRICEDISC(B1,B2,5.25%,100,2)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);

            }
        }

        [TestMethod]
        public void PriceDiscShouldReturnCorrectResult2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with correct input values");
                sheet.Cells["B1"].Value = new System.DateTime(2008, 2, 16);
                sheet.Cells["B2"].Value = new System.DateTime(2018, 5, 16);
                sheet.Cells["A1"].Formula = "PRICEDISC(B1,B2,50%,2000,4)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 0);
                Assert.AreEqual(-8250d, result);

            }
        }
    }

}
