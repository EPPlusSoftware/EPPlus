using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Finance
{
    [TestClass]
    public class ReceivedTest : TestBase
    {
        [TestMethod]
        public void ReceivedTestWithRegularInput()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with regular input: ");
                sheet.Cells["A1"].Value = new System.DateTime(2008, 2, 15);
                sheet.Cells["A2"].Value = new System.DateTime(2008, 4, 13);
                sheet.Cells["A4"].Formula = "RECEIVED(A1, A2, 180, 6.10%, 4)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A4"].Value, 7);
                Assert.AreEqual(181.7865579d, result);
            }
        }

        [TestMethod]
        public void ReceivedIncorrectDiscount()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect discount: ");
                sheet.Cells["A1"].Value = new System.DateTime(2008, 2, 15);
                sheet.Cells["A2"].Value = new System.DateTime(2008, 4, 13);
                sheet.Cells["A4"].Formula = "RECEIVED(A1, A2, 180, 0%, 4)";
                sheet.Calculate();
                var result = sheet.Cells["A4"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void ReceivedIncorrectInvestmentParameter()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect investments: ");
                sheet.Cells["A1"].Value = new System.DateTime(2008, 2, 15);
                sheet.Cells["A2"].Value = new System.DateTime(2008, 4, 13);
                sheet.Cells["A4"].Formula = "RECEIVED(A1, A2, 0, 3%, 4)";
                sheet.Calculate();
                var result = sheet.Cells["A4"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void ReceivedIncorrectDates()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect dates: ");
                sheet.Cells["A1"].Value = new System.DateTime(2008, 2, 15);
                sheet.Cells["A2"].Value = new System.DateTime(2008, 4, 13);
                sheet.Cells["A4"].Formula = "RECEIVED(A2, A1, 180, 3%, 4)";
                sheet.Calculate();
                var result = sheet.Cells["A4"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void ReceivedRandomTest()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with regular input: ");
                sheet.Cells["A1"].Value = new System.DateTime(2008, 2, 15);
                sheet.Cells["A2"].Value = new System.DateTime(2008, 4, 13);
                sheet.Cells["A4"].Formula = "RECEIVED(A1, A2, 80.324, 70.43%, 2)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A4"].Value, 8);
                Assert.AreEqual(90.60499965d, result);
            }
        }
    }
}
