using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class GammainvTests
    {

        [TestMethod]
        public void GammaInvShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 0.068094;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 2;

                sheet.Cells["B5"].Formula = "GAMMA.INV(A2,A3,A4)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 7);
                Assert.AreEqual(10.0000112, result);
            }
        }

        [TestMethod]
        public void GammaInvShouldReturnCorrectResult2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 0.068094;
                sheet.Cells["A3"].Value = 1;
                sheet.Cells["A4"].Value = 1;

                sheet.Cells["B5"].Formula = "GAMMA.INV(A2,A3,A4)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 7);
                Assert.AreEqual(0.0705233, result);
            }
        }

        [TestMethod]
        public void GammaInvShouldReturnCorrectDecimals()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 0.070523328;
                sheet.Cells["A3"].Value = 12.231321;
                sheet.Cells["A4"].Value = 1.42332;

                sheet.Cells["B5"].Formula = "GAMMA.INV(A2,A3,A4)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 7);
                Assert.AreEqual(10.7138623, result);
            }
        }

        [TestMethod]
        public void GammaInvShouldReturnNumError()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 0.070523328;
                sheet.Cells["A3"].Value = 12.231321;
                sheet.Cells["A4"].Value = -1.42332;

                sheet.Cells["B5"].Formula = "GAMMA.INV(A2,A3,A4)";
                sheet.Calculate();

                var result =sheet.Cells["B5"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void GammaInvShouldZero()
        {
            using (var package = new ExcelPackage())
            {
                var sheet=package.Workbook.Worksheets.Add("test"); 
                sheet.Cells["A3"].Value = 1;
                sheet.Cells["A4"].Value = 1;

                sheet.Cells["B5"].Formula = "GAMMA.INV(A2,A3,A4)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 7);
                Assert.AreEqual(0, result);
            }
        }

        [TestMethod]
        public void GammaInvArrayTest()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;

                sheet.Cells["A5"].Value = 1;
                sheet.Cells["A6"].Value = 1;
                sheet.Cells["A7"].Value = 0.068094;

                sheet.Cells["B5"].Formula = "GAMMA.INV(A7,A5:A6,A2:A3)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 7);
                var result2 = System.Math.Round((double)sheet.Cells["B6"].Value, 7);
                Assert.AreEqual(0.1410467, result);
                Assert.AreEqual(0.21157, result2);

            }
        }

        [TestMethod]
        public void GammaInvArrayTest2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;

                sheet.Cells["A5"].Value = 1;
                sheet.Cells["A6"].Value = 1;
                sheet.Cells["B5"].Value = 4.24234;
                sheet.Cells["B6"].Value = 10;

                sheet.Cells["A7"].Value = 0.068094;
                sheet.Cells["A8"].Value = 0.56345;

                sheet.Cells["B8"].Formula = "GAMMA.INV(A7:A8,A5:B6,A2:A3)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B8"].Value, 7);
                var result2 = System.Math.Round((double)sheet.Cells["B9"].Value, 7);
                var result3 = System.Math.Round((double)sheet.Cells["C8"].Value, 7);
                var result4 = System.Math.Round((double)sheet.Cells["C9"].Value, 7);
                Assert.AreEqual(0.1410467, result);
                Assert.AreEqual(2.4865571, result2);
                Assert.AreEqual(3.3405934, result3);
                Assert.AreEqual(30.5173225, result4);
            }
        }
    }
}