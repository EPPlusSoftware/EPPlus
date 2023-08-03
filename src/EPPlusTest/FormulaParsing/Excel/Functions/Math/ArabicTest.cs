using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
    [TestClass]
    public class ArabicTest : TestBase
    {

        [TestMethod]
        public void ArabicTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with classical roman style: ");
                sheet.Cells["A1"].Formula = "=ARABIC(\"CDXCIX\")";
                sheet.Calculate();
                var result = (double)sheet.Cells["A1"].Value;
                Assert.AreEqual(499d, result);
            }
        }

        [TestMethod]
        public void ArabicTest2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with more cocise version: ");
                sheet.Cells["A1"].Formula = "=ARABIC(\"LDVLIV\")";
                sheet.Calculate();
                var result = (double)sheet.Cells["A1"].Value;
                Assert.AreEqual(499d, result);
            }
        }

        [TestMethod]
        public void ArabicTest3()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with more cocise version: ");
                sheet.Cells["A1"].Formula = "=ARABIC(\"XDIX\")";
                sheet.Calculate();
                var result = (double)sheet.Cells["A1"].Value;
                Assert.AreEqual(499d, result);
            }
        }

        [TestMethod]
        public void ArabicTest4()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with more cocise version: ");
                sheet.Cells["A1"].Formula = "=ARABIC(\"VDIV\")";
                sheet.Calculate();
                var result = (double)sheet.Cells["A1"].Value;
                Assert.AreEqual(499d, result);
            }
        }

        [TestMethod]
        public void ArabicTest5()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Simplified version: ");
                sheet.Cells["A1"].Formula = "=ARABIC(\"ID\")";
                sheet.Calculate();
                var result = (double)sheet.Cells["A1"].Value;
                Assert.AreEqual(499d, result);
            }
        }

        [TestMethod]
        public void ArabicTest6()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Weird input: ");
                sheet.Cells["A1"].Formula = "=ARABIC(\"CDLIDCDLDVLIVXIXLVD\")";
                sheet.Calculate();
                var result = (double)sheet.Cells["A1"].Value;
                Assert.AreEqual(2062d, result);
            }
        }

        [TestMethod]
        public void ArabicTest7()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Weird input: ");
                sheet.Cells["A1"].Formula = "=ARABIC(\"MLIXDCM\")";
                sheet.Calculate();
                var result = (double)sheet.Cells["A1"].Value;
                Assert.AreEqual(1339d, result);
            }
        }

        [TestMethod]
        public void ArabicTestInvalidInput()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Invalid input: ");
                sheet.Cells["A1"].Formula = "=ARABIC(\"MLI6jhXDCM\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result);
            }
        }

        [TestMethod]
        public void ArabicTestEmptyString()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Empty string: ");
                sheet.Cells["A1"].Formula = "=ARABIC(\"\")";
                sheet.Calculate();
                var result = (double)sheet.Cells["A1"].Value;
                Assert.AreEqual(0d, result);
            }
        }

        [TestMethod]
        public void ArabicTestCaseInsensetivity()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Weird input: ");
                sheet.Cells["A1"].Formula = "=ARABIC(\"MLixDCm\")";
                sheet.Calculate();
                var result = (double)sheet.Cells["A1"].Value;
                Assert.AreEqual(1339d, result);
            }
        }

        [TestMethod]
        public void ArabicTestLeadingAndTrailingSpace()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Weird input: ");
                sheet.Cells["A1"].Formula = "=ARABIC(\"     MLIXDCM    \")";
                sheet.Calculate();
                var result = (double)sheet.Cells["A1"].Value;
                Assert.AreEqual(1339d, result);
            }
        }

        [TestMethod]
        public void ArabicTestError()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Weird input: ");
                sheet.Cells["A1"].Formula = "=ARABIC(\"     MLIX    DCM    \")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result);
            }
        }

        [TestMethod]
        public void ArabicTestNegativeRoman()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Weird input: ");
                sheet.Cells["A1"].Formula = "=ARABIC(\"-MLIXDCM\")";
                sheet.Calculate();
                var result = (double)sheet.Cells["A1"].Value;
                Assert.AreEqual(-1339d, result);
            }
        }

        [TestMethod]
        public void ArabicRangeTest()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with more cocise version: ");
                sheet.Cells["A5"].Value = "VDIV";
                sheet.Cells["A6"].Value = "XXI";
                sheet.Cells["A1"].Formula = "=ARABIC(A5:A6)";
                sheet.Calculate();
                var result = (double)sheet.Cells["A1"].Value;
                var result2 = (double)sheet.Cells["A2"].Value;
                Assert.AreEqual(499d, result);
                Assert.AreEqual(21d, result2);
            }
        }

    }
}
