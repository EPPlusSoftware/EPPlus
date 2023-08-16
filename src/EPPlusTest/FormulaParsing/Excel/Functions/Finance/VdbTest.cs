using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Finance
{
    [TestClass]
    public class VdbTest : TestBase
    {

        [TestMethod]
        public void VdbTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct results: ");
                sheet.Cells["A1"].Formula = "VDB(2400,300,10,0,1,3,TRUE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 2);
                Assert.AreEqual(720.00d, result);
            }
        }

        [TestMethod]
        public void VdbTestNoSwitchIsFalse()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct results: ");
                sheet.Cells["A1"].Formula = "VDB(2000,90,50,8,32,2,FALSE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 2);
                Assert.AreEqual(905.18d, result);
            }
        }

        [TestMethod]
        public void VdbTestNoSwitchIsTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct results: ");
                sheet.Cells["A1"].Formula = "VDB(2000, 90, 50, 8, 32, 2, TRUE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 2);
                Assert.AreEqual(901.14d, result);
            }
        }

        [TestMethod]
        public void VdbTestNoSwitchOmitted()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct results: ");
                sheet.Cells["A1"].Formula = "VDB(2000, 90, 50, 8, 32, 2)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 2);
                Assert.AreEqual(905.18d, result);
            }
        }

        [TestMethod]
        public void VdbTestIncorrectCost()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct results: ");
                sheet.Cells["A1"].Formula = "VDB(-2000, 90, 50, 8, 32, 2)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void VdbTestIncorrectSalvage()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct results: ");
                sheet.Cells["A1"].Formula = "VDB(2000, -90, 50, 8, 32, 2)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void VdbTestIncorrectPeriod()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct results: ");
                sheet.Cells["A1"].Formula = "VDB(2000, 90, 1, 8, 32, 2)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void VdbTestEndLessThanStartPeriod()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct results: ");
                sheet.Cells["A1"].Formula = "VDB(2000, 90, 50, 32, 8, 2)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void VdbTestUnevenPeriodNoSwitchTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct results: ");
                sheet.Cells["A1"].Formula = "VDB(2000, 90, 50, 8, 32.6, 2, TRUE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 2);
                Assert.AreEqual(914.14d, result);
            }
        }

        [TestMethod]
        public void VdbTestUnevenPeriodNoSwitchFalse()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct results: ");
                sheet.Cells["A1"].Formula = "VDB(2000, 90, 50, 8, 32.6, 2, FALSE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 2);
                Assert.AreEqual(920.10d, result);
            }
        }

        [TestMethod]
        public void VdbTestIncorrectFactor()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct results: ");
                sheet.Cells["A1"].Formula = "VDB(2000, 90, 50, 8, 32, -2)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]

        public void VdbTestDayDepreciation()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct result: ");
                sheet.Cells["A1"].Formula = "VDB(2400, 300, 3650, 0, 1)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 2);
                Assert.AreEqual(1.32d, result);
            }
        }

        [TestMethod]

        public void VdbTestMonthDepreciation()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct result: ");
                sheet.Cells["A1"].Formula = "VDB(2400, 300, 120, 0, 1)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 2);
                Assert.AreEqual(40.00d, result);
            }
        }

        [TestMethod]
        public void VdbFinalTest()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct result: ");
                sheet.Cells["A1"].Formula = "VDB(2378.39, 346.3554, 120, 6, 18, 2.9843, FALSE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 2);
                Assert.AreEqual(533.32d, result);
            }
        }

    }
}
