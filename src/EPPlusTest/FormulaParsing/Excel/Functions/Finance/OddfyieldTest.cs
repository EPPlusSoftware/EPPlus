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
        public void OddfyieldCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with correct input values");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 2, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 2, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 1);
                sheet.Cells["B4"].Value = new System.DateTime(2019, 2, 15);
                sheet.Cells["A1"].Formula = "ODDFYIELD(B1,B2,B3,B4,1%,60,100,4,0)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.187138323, result);
            }
        }

        [TestMethod]
        public void OddfyieldWithExtremeInput()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with very long maturity and high interest rate: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 2, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2032, 2, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 1);
                sheet.Cells["B4"].Value = new System.DateTime(2019, 2, 15);
                sheet.Cells["A1"].Formula = "ODDFYIELD(B1,B2,B3,B4,180%,60,100,1,0)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(3.003480844, result);
            }
        }

        [TestMethod]
        public void OddfyieldShortPeriodTest()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with short period");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 02, 01);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 02, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 02, 01);
                sheet.Cells["B4"].Value = new System.DateTime(2019, 02, 15);
                sheet.Cells["A1"].Formula = "ODDFYIELD(B1,B2,B3,B4,1%,120,100,2,0)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(-0.050099877, result);


            }
        }

        [TestMethod]
        public void OddfyieldLongPeriodTest()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with long period:");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 02, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 02, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 01);
                sheet.Cells["B4"].Value = new System.DateTime(2019, 02, 15);
                sheet.Cells["A1"].Formula = "ODDFYIELD(B1,B2,B3,B4,1%,120,100,4,0)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(-0.051190686, result);
            }
        }

        [TestMethod]
        public void OddfyieldInvalidBasis()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with invalid basis: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 02, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 02, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 01);
                sheet.Cells["B4"].Value = new System.DateTime(2019, 02, 15);
                sheet.Cells["A1"].Formula = "ODDFYIELD(B1,B2,B3,B4,1%,120,100,4,8)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }

        }

        [TestMethod]
        public void OddfyieldIncorrectRate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with negative rate: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 02, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 02, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 01);
                sheet.Cells["B4"].Value = new System.DateTime(2019, 02, 15);
                sheet.Cells["A1"].Formula = "ODDFYIELD(B1,B2,B3,B4,-5%,120,100,4,0)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void OddfyieldInvalidDates()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return error when dates are incorrect: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 02, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2017, 02, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 01);
                sheet.Cells["B4"].Value = new System.DateTime(2019, 02, 15);
                sheet.Cells["A1"].Formula = "ODDFYIELD(B1,B2,B3,B4,1%,150,100,4,1)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void OddfyieldIncorrectFrequency()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect frequency: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 02, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 02, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 01);
                sheet.Cells["B4"].Value = new System.DateTime(2019, 02, 15);
                sheet.Cells["A1"].Formula = "ODDFYIELD(B1,B2,B3,B4,1%,100,100,5,0)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void OddfyieldNoBasisArgument()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with no basis argument: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 02, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 02, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 01);
                sheet.Cells["B4"].Value = new System.DateTime(2019, 02, 15);
                sheet.Cells["A1"].Formula = "ODDFYIELD(B1,B2,B3,B4,1%,10,100,1)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(1.209580503, result);
            }
        }

        [TestMethod]
        public void OddfyieldZeroOrNegativePrice()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Should return num error when price <= 0: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 02, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 02, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 01);
                sheet.Cells["B4"].Value = new System.DateTime(2019, 02, 15);
                sheet.Cells["A1"].Formula = "ODDFYIELD(B1,B2,B3,B4,1%,-50,100,2,0)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }
    }
}
