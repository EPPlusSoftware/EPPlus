using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class LogNormDotDistTests
    {

        [TestMethod]
        public void LogNormDotDistShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 4;
                sheet.Cells["A3"].Value = 3.5;
                sheet.Cells["A4"].Value = 1.2;

                sheet.Cells["A5"].Formula = "LOGNORM.DIST(A2,A3,A4,TRUE)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["A5"].Value, 7);
                Assert.AreEqual(0.0390836, result);
            }
        }

        [TestMethod]
        public void LogNormDotDistShouldReturnCorrectResultWhenFalse()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 4;
                sheet.Cells["A3"].Value = 3.5;
                sheet.Cells["A4"].Value = 1.2;

                sheet.Cells["A5"].Formula = "LOGNORM.DIST(A2,A3,A4,FALSE)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["A5"].Value, 7);
                Assert.AreEqual(0.0176176, result);
            }
        }

        [TestMethod]
        public void LogNormDotDistShouldReturnCorrectResultArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = 4;
                sheet.Cells["A4"].Value = 3.5;
                sheet.Cells["A5"].Value = 1.2;

                sheet.Cells["A6"].Formula = "LOGNORM.DIST(A2:A3,A4,A5, TRUE)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["A6"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["A7"].Value, 9);
                Assert.AreEqual(0.022687768, result);
                Assert.AreEqual(0.039083556, result2);
            }
        }

        [TestMethod]
        public void LogNormDotDistShouldReturnCorrectResultArray2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = 4;
                sheet.Cells["A4"].Value = 3.5;
                sheet.Cells["A5"].Value = 3;
                sheet.Cells["A6"].Value = 1.2;

                sheet.Cells["B7"].Formula = "LOGNORM.DIST(A2:A3,A4:A5,A6, TRUE)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B7"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                Assert.AreEqual(0.022687768, result);
                Assert.AreEqual(0.089352232, result2);
            }
        }

        [TestMethod]
        public void LogNormDotDistShouldReturnCorrectResultArray3()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = 4;
                sheet.Cells["A4"].Value = 3.5;
                sheet.Cells["A5"].Value = 3;
                sheet.Cells["A6"].Value = 1.2;

                sheet.Cells["B7"].Formula = "LOGNORM.DIST(A2:A4,A4:A5,A6, TRUE)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B7"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                var result3 = sheet.Cells["B9"].Value;
                Assert.AreEqual(0.022687768, result); 
                Assert.AreEqual(0.089352232, result2);
                Assert.AreEqual(ErrorValues.NAError, result3);
            }
        }

        [TestMethod]
        public void LogNormDotDistShouldReturnCorrectResultArray4()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = 4;
                sheet.Cells["A4"].Value = 3.5;
                sheet.Cells["A5"].Value = 3;
                sheet.Cells["A6"].Value = 1.2;
                sheet.Cells["A7"].Value = "TRUE";

                sheet.Cells["B7"].Formula = "LOGNORM.DIST(A2:A4,A4:A5,A6,A7)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B7"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                var result3 = sheet.Cells["B9"].Value;
                Assert.AreEqual(0.022687768, result);
                Assert.AreEqual(0.089352232, result2);
            }
        }

        [TestMethod]
        public void LogNormDotDistShouldReturnCorrectResultArray5()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = 4;
                sheet.Cells["A4"].Value = 3.5;
                sheet.Cells["A5"].Value = 3;
                sheet.Cells["A6"].Value = 1.2;
                sheet.Cells["A7"].Value = false;

                sheet.Cells["B7"].Formula = "LOGNORM.DIST(A2:A4,A4:A5,A6,A7)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B7"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                Assert.AreEqual(0.014962834, result);
                Assert.AreEqual(0.033650174, result2);
            }
        }

        [TestMethod]
        public void LogNormDotDistShouldReturnCorrectResultArray6()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = 4;
                sheet.Cells["A4"].Value = 3.5;
                sheet.Cells["A5"].Value = 3;
                sheet.Cells["A6"].Value = 1.2;
                sheet.Cells["A7"].Value = false;

                sheet.Cells["B7"].Formula = "LOGNORM.DIST(A2:A4,A4,A5:A6,A7)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B7"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                Assert.AreEqual(0.032176038, result);
                Assert.AreEqual(0.017617597, result2);
            }
        }

        [TestMethod]
        public void LogNormDotDistShouldReturnCorrectResultArray7()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = 4;
                sheet.Cells["A4"].Value = 3.5;
                sheet.Cells["A5"].Value = 3;
                sheet.Cells["A6"].Value = 1.2;
                sheet.Cells["A7"].Value = true;

                sheet.Cells["B7"].Formula = "LOGNORM.DIST(A2:A4,A4:A5,A6,A7)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B7"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                var result3 = sheet.Cells["B9"].Value;
                Assert.AreEqual(0.022687768, result);
                Assert.AreEqual(0.089352232, result2);
            }
        }

        [TestMethod]
        public void LogNormDotDistShouldReturnCorrectResultArray8()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = 4;
                sheet.Cells["A4"].Value = 3.5;
                sheet.Cells["A5"].Value = 3;
                sheet.Cells["A6"].Value = 1.2;
                sheet.Cells["A7"].Value ="FALSE";

                sheet.Cells["B7"].Formula = "LOGNORM.DIST(A2:A4,A4,A5:A6,A7)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B7"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                Assert.AreEqual(0.032176038, result);
                Assert.AreEqual(0.017617597, result2);
            }
        }

        [TestMethod]
        public void LogNormDotDistShouldReturnCorrectResultArray9()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = 4;
                sheet.Cells["A4"].Value = 3.5;
                sheet.Cells["A5"].Value = 3;
                sheet.Cells["A6"].Value = 1.2;
                sheet.Cells["A7"].Value = true;
                sheet.Cells["A8"].Value = false;

                sheet.Cells["B7"].Formula = "LOGNORM.DIST(A2:A3,A4,A5:A6,A7:A8)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B7"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                var result3 = sheet.Cells["B9"].Value;
                Assert.AreEqual(0.211721421, result);
                Assert.AreEqual(0.017617597, result2);
            }
        }
    }
}
