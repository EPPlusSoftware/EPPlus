using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class ZDotTestTests
    {
        [TestMethod]
        public void ZDotTestTest1()
        {
            using (var package = new ExcelPackage())
            {
                ExcelWorksheet sheet = LoadSheet(package);
                sheet.Cells["E12"].Formula = "Z.TEST(E2:E11,4)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["E12"].Value, 12);
                Assert.AreEqual(0.090574196851, result);
            }
        }

        [TestMethod]
        public void ZDotTestShouldReturnWithSigmaParameter()
        {
            using (var package = new ExcelPackage())
            {
                ExcelWorksheet sheet = LoadSheet(package);

                sheet.Cells["E12"].Formula = "Z.TEST(E2:E11,4,3)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["E12"].Value, 12);
                Assert.AreEqual(0.123125849846, result);
            }
        }

        [TestMethod]
        public void ZDotTestShouldReturnCorrectResultEmptyCells()
        {
            using (var package = new ExcelPackage())
            {
                ExcelWorksheet sheet = LoadSheet(package);
                sheet.Cells["E12"].Formula = "Z.TEST(E2:G11,4,3)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["E12"].Value, 12);
                Assert.AreEqual(0.123125849846, result);
            }
        }

        [TestMethod]
        public void ZDotTestShouldReturnError()
        {
            using (var package = new ExcelPackage())
            {
                ExcelWorksheet sheet = LoadSheet(package);
                sheet.Cells["E12"].Formula = "Z.TEST(E2:G11,4,0)";
                sheet.Calculate();

                var result = sheet.Cells["E12"].Value;
                Assert.AreEqual(ErrorValues.NumError, result);
            }
        }

        [TestMethod]
        public void ZDotTestShouldReturnCorrectResult2()
        {
            using (var package = new ExcelPackage())
            {
                ExcelWorksheet sheet = LoadSheet(package);
                sheet.Cells["E12"].Formula = "Z.TEST(E2:G11,0,3)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["E12"].Value, 12);
                Assert.AreEqual(0.000000038106, result);
            }
        }

        [TestMethod]
        public void ZDotTestShouldReturnCorrectResult3()
        {
            using (var package = new ExcelPackage())
            {
                ExcelWorksheet sheet = LoadSheet(package);
                sheet.Cells["E12"].Formula = "Z.TEST(E2:G11,-1,3)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["E12"].Value, 12);
                Assert.AreEqual(0.000000000064, result);
            }
        }


        [TestMethod]
        public void ZDotTestShouldReturnError2()
        {
            using (var package = new ExcelPackage())
            {
                ExcelWorksheet sheet = LoadSheet(package);
                sheet.Cells["E12"].Formula = "Z.TEST(E2:G11,-1,3,4)";
                sheet.Calculate();

                var result = sheet.Cells["E12"].Value;
                Assert.AreEqual(ErrorValues.ValueError, result);
            }
        }

        [TestMethod]
        public void ZDotTestShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                ExcelWorksheet sheet = LoadSheet(package);
                sheet.Cells["E12"].Formula = "Z.TEST(E2:G11,3.67342,4.23432)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["E12"].Value, 12);
                Assert.AreEqual(0.143347609979, result);
            }
        }

        private static ExcelWorksheet LoadSheet(ExcelPackage package)
        {
            var sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["E2"].Value = 3;
            sheet.Cells["E3"].Value = 6;
            sheet.Cells["E4"].Value = 7;
            sheet.Cells["E5"].Value = 8;
            sheet.Cells["E6"].Value = 6;
            sheet.Cells["E7"].Value = 5;
            sheet.Cells["E8"].Value = 4;
            sheet.Cells["E9"].Value = 2;
            sheet.Cells["E10"].Value = 1;
            sheet.Cells["E11"].Value = 9;
            return sheet;
        }
    }
}
