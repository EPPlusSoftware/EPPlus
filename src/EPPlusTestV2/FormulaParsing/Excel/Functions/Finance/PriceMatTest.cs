using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Finance
{
    [TestClass]
    public class PriceMatTest : TestBase
    {
        [TestMethod]
        public void PriceMatTestRegularInput()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with regular input: ");
                sheet.Cells["A1"].Value = new System.DateTime(2008, 2, 15);
                sheet.Cells["A2"].Value = new System.DateTime(2008, 4, 13);
                sheet.Cells["A3"].Value = new System.DateTime(2007, 11, 11);
                sheet.Cells["A4"].Formula = "PRICEMAT(A1, A2, A3, 6.10%, 6.10%, 0)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A4"].Value, 8);
                Assert.AreEqual(99.98449888d, result);
            }
        }

            
        [TestMethod]
        public void PriceMatIncorrectRate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect rate: ");
                sheet.Cells["A1"].Value = new System.DateTime(2008, 2, 15);
                sheet.Cells["A2"].Value = new System.DateTime(2008, 4, 13);
                sheet.Cells["A3"].Value = new System.DateTime(2007, 11, 11);
                sheet.Cells["A4"].Formula = "PRICEMAT(A1, A2, A3, -6.10%, 6.10%, 0)";
                sheet.Calculate();
                //var result = System.Math.Round((double)sheet.Cells["A4"].Value, 8);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), sheet.Cells["A4"].Value);
            }
        }

        [TestMethod]
        public void PriceMatIncorrectYield()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect yield: ");
                sheet.Cells["A1"].Value = new System.DateTime(2008, 2, 15);
                sheet.Cells["A2"].Value = new System.DateTime(2008, 4, 13);
                sheet.Cells["A3"].Value = new System.DateTime(2007, 11, 11);
                sheet.Cells["A4"].Formula = "PRICEMAT(A1, A2, A3, 6.10%, -6.10%, 0)";
                sheet.Calculate();
                //var result = System.Math.Round((double)sheet.Cells["A4"].Value, 8);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), sheet.Cells["A4"].Value);
            }
        }

        [TestMethod]
        public void PriceMatIncorrectDates()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect dates: ");
                sheet.Cells["A1"].Value = new System.DateTime(2008, 2, 15);
                sheet.Cells["A2"].Value = new System.DateTime(2008, 4, 13);
                sheet.Cells["A3"].Value = new System.DateTime(2007, 11, 11);
                sheet.Cells["A4"].Formula = "PRICEMAT(A2, A1, A3, 6.10%, 6.10%, 0)";
                sheet.Calculate();
                //var result = System.Math.Round((double)sheet.Cells["A4"].Value, 8);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), sheet.Cells["A4"].Value);
            }
        }

        [TestMethod]
        public void PriceMatFinalTest()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with regular input: ");
                sheet.Cells["A1"].Value = new System.DateTime(2015, 2, 9);
                sheet.Cells["A2"].Value = new System.DateTime(2070, 4, 13);
                sheet.Cells["A3"].Value = new System.DateTime(2007, 11, 11);
                sheet.Cells["A4"].Formula = "PRICEMAT(A1, A2, A3, 15.97%, 100%, 4)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A4"].Value, 8);
                Assert.AreEqual(-96.16856813d, result);
            }
        }

    }
}
