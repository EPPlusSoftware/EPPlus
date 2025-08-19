using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class ProbTests
    {
        [TestMethod]
        public void ProbShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A2"].Value = 0;
                sheet.Cells["A3"].Value = 1;
                sheet.Cells["A4"].Value = 2;
                sheet.Cells["A5"].Value = 3;

                sheet.Cells["B2"].Value = 0.2;
                sheet.Cells["B3"].Value = 0.3;
                sheet.Cells["B4"].Value = 0.1;
                sheet.Cells["B5"].Value = 0.4;

                sheet.Cells["C2"].Value = 1;
                sheet.Cells["D2"].Value = 3;

                sheet.Cells["E9"].Formula = "PROB(A2:A5,B2:B5,C2,D2)";
                sheet.Calculate();

                var result = sheet.Cells["E9"].Value;
                Assert.AreEqual(0.8, result);
            }
        }

        [TestMethod]
        public void ProbShouldReturnCorrectResult2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A2"].Value = 0;
                sheet.Cells["A3"].Value = 1;
                sheet.Cells["A4"].Value = 2;
                sheet.Cells["A5"].Value = 3;

                sheet.Cells["B2"].Value = 0.2;
                sheet.Cells["B3"].Value = 0.3;
                sheet.Cells["B4"].Value = 0.1;
                sheet.Cells["B5"].Value = 0.4;

                sheet.Cells["C2"].Value = 2;

                sheet.Cells["E9"].Formula = "PROB(A2:A5,B2:B5,C2)";
                sheet.Calculate();

                var result = sheet.Cells["E9"].Value;
                Assert.AreEqual(0.1, result);
            }
        }

        [TestMethod]
        public void ProbShouldReturnCorrectResult3()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["E9"].Formula = "PROB(1,1,1)";
                sheet.Calculate();

                var result = sheet.Cells["E9"].Value;
                Assert.AreEqual(1d, result);
            }
        }

        [TestMethod]
        public void ProbShouldReturnCorrectResult4()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["B2"].Value = 0.2;
                sheet.Cells["B3"].Value = 0.3;
                sheet.Cells["B4"].Value = 0.1;
                sheet.Cells["B5"].Value = 0.4;

                sheet.Cells["C2"].Value = 2;

                sheet.Cells["E9"].Formula = "PROB({0,1,2,3},B2:B5,C2)";
                sheet.Calculate();

                var result = sheet.Cells["E9"].Value;
                Assert.AreEqual(0.1, result);
            }
        }

        [TestMethod]
        public void ProbShouldReturnError()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["B2"].Value = 0.3;
                sheet.Cells["B3"].Value = 0.3;
                sheet.Cells["B4"].Value = 0.1;
                sheet.Cells["B5"].Value = 0.4;

                sheet.Cells["C2"].Value = 2;

                sheet.Cells["E9"].Formula = "PROB({0,1,2,3},B2:B5,C2)";
                sheet.Calculate();

                var result = sheet.Cells["E9"].Value;
                Assert.AreEqual(ErrorValues.NumError, result);
            }
        }

        [TestMethod]
        public void ProbShouldReturnError2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["B2"].Value = 0.2;
                sheet.Cells["B3"].Value = 0.3;
                sheet.Cells["B4"].Value = 0.1;
                sheet.Cells["B5"].Value = 0.4;

                sheet.Cells["C2"].Value = 2;

                sheet.Cells["E9"].Formula = "PROB({0,1,2},B2:B5,C2)";
                sheet.Calculate();

                var result = sheet.Cells["E9"].Value;
                Assert.AreEqual(ErrorValues.NAError, result);
            }
        }


        [TestMethod]
        public void ProbShouldReturnCorrectResult5()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["B2"].Value = 0.25;
                sheet.Cells["B3"].Value = 0.3;
                sheet.Cells["B4"].Value = 0.1;
                sheet.Cells["B5"].Value = 0.35;

                sheet.Cells["C2"].Value = 2;

                sheet.Cells["E9"].Formula = "PROB({0,1,2,3},B2:B5,C2)";
                sheet.Calculate();

                var result = sheet.Cells["E9"].Value;
                Assert.AreEqual(0.1, result);
            }
        }

        [TestMethod]
        public void ProbShouldReturnCorrectResultWhenLlandHlisNull()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["B2"].Value = 0.25;
                sheet.Cells["B3"].Value = 0.3;
                sheet.Cells["B4"].Value = 0.1;
                sheet.Cells["B5"].Value = 0.35;

                sheet.Cells["E9"].Formula = "PROB({0,1,2,3},B2:B5,C2,D2)";
                sheet.Calculate();

                var result = sheet.Cells["E9"].Value;
                Assert.AreEqual(0.25, result);
            }
        }
    }
}
