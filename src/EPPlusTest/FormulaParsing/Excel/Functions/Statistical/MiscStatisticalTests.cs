using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class MiscStatisticalTests
    {
        [TestMethod]
        public void CorrelTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 3;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 4;
                sheet.Cells["A4"].Value = 5;
                sheet.Cells["A5"].Value = 6;
                sheet.Cells["B1"].Value = 9;
                sheet.Cells["B2"].Value = 7;
                sheet.Cells["B3"].Value = 12;
                sheet.Cells["B4"].Value = 15;
                sheet.Cells["B5"].Value = 17;
                sheet.Cells["B6"].Formula = "CORREL(A1:A5,B1:B5)";
                sheet.Calculate();
                var result = sheet.Cells["B6"].Value;

                Assert.AreEqual(0.997054, System.Math.Round((double)result, 6));
            }
        }

        [TestMethod]
        public void FisherTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 0.75;
                sheet.Cells["A2"].Formula = "FISHER(A1)";
                sheet.Calculate();
                var result = sheet.Cells["A2"].Value;

                Assert.AreEqual(0.9729551, System.Math.Round((double)result, 7));
            }
        }

        [TestMethod]
        public void FisherInvTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 0.9729551;
                sheet.Cells["A2"].Formula = "FISHERINV(A1)";
                sheet.Calculate();
                var result = sheet.Cells["A2"].Value;

                Assert.AreEqual(0.75, System.Math.Round((double)result, 2));
            }
        }

        [TestMethod]
        public void GeomeanTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 4;
                sheet.Cells["A2"].Value = 5;
                sheet.Cells["A3"].Value = 8;
                sheet.Cells["A4"].Value = 7;
                sheet.Cells["A5"].Value = 11;
                sheet.Cells["A6"].Value = 4;
                sheet.Cells["A7"].Value = 3;
                sheet.Cells["B6"].Formula = "GEOMEAN(A1:A7)";
                sheet.Calculate();
                var result = sheet.Cells["B6"].Value;

                Assert.AreEqual(5.476987, System.Math.Round((double)result, 6));
            }
        }

        [TestMethod]
        public void HarmeanTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 4;
                sheet.Cells["A2"].Value = 5;
                sheet.Cells["A3"].Value = 8;
                sheet.Cells["A4"].Value = 7;
                sheet.Cells["A5"].Value = 11;
                sheet.Cells["A6"].Value = 4;
                sheet.Cells["A7"].Value = 3;
                sheet.Cells["B6"].Formula = "HARMEAN(A1:A7)";
                sheet.Calculate();
                var result = sheet.Cells["B6"].Value;

                Assert.AreEqual(5.028376, System.Math.Round((double)result, 6));
            }
        }

        [TestMethod]
        public void PearsonTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 9;
                sheet.Cells["A2"].Value = 7;
                sheet.Cells["A3"].Value = 5;
                sheet.Cells["A4"].Value = 3;
                sheet.Cells["A5"].Value = 1;
                sheet.Cells["B1"].Value = 10;
                sheet.Cells["B2"].Value = 6;
                sheet.Cells["B3"].Value = 1;
                sheet.Cells["B4"].Value = 5;
                sheet.Cells["B5"].Value = 3;
                sheet.Cells["B6"].Formula = "PEARSON(A1:A5,B1:B5)";
                sheet.Calculate();
                var result = sheet.Cells["B6"].Value;

                Assert.AreEqual(0.699379, System.Math.Round((double)result, 6));
            }
        }

        [TestMethod]
        public void RsqTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 1;
                sheet.Cells["A5"].Value = 8;
                sheet.Cells["A6"].Value = 7;
                sheet.Cells["A7"].Value = 5;
                sheet.Cells["B1"].Value = 6;
                sheet.Cells["B2"].Value = 5;
                sheet.Cells["B3"].Value = 11;
                sheet.Cells["B4"].Value = 7;
                sheet.Cells["B5"].Value = 5;
                sheet.Cells["B6"].Value = 4;
                sheet.Cells["B7"].Value = 4;
                sheet.Cells["B8"].Formula = "RSQ(A1:A7,B1:B7)";
                sheet.Calculate();
                var result = sheet.Cells["B8"].Value;

                Assert.AreEqual(0.05795, System.Math.Round((double)result, 5));
            }
        }

        [TestMethod]
        public void SkewTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 3;
                sheet.Cells["A2"].Value = 4;
                sheet.Cells["A3"].Value = 5;
                sheet.Cells["A4"].Value = 2;
                sheet.Cells["A5"].Value = 3;
                sheet.Cells["A6"].Value = 4;
                sheet.Cells["A7"].Value = 5;
                sheet.Cells["A8"].Value = 6;
                sheet.Cells["A9"].Value = 4;
                sheet.Cells["A10"].Value = 7;
                sheet.Cells["B8"].Formula = "SKEW(A1:A10)";
                sheet.Calculate();
                var result = sheet.Cells["B8"].Value;

                Assert.AreEqual(0.359543, System.Math.Round((double)result, 6));
            }
        }

        [TestMethod]
        public void StandardizeTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 42;
                sheet.Cells["A2"].Value = 40;
                sheet.Cells["A3"].Value = 1.5;
                sheet.Cells["B8"].Formula = "STANDARDIZE(A1,A2,A3)";
                sheet.Calculate();
                var result = sheet.Cells["B8"].Value;

                Assert.AreEqual(1.333333, System.Math.Round((double)result, 6));
            }
        }

        [TestMethod]
        public void ForecastTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 6;
                sheet.Cells["A2"].Value = 7;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 15;
                sheet.Cells["A5"].Value = 21;
                sheet.Cells["B1"].Value = 20;
                sheet.Cells["B2"].Value = 28;
                sheet.Cells["B3"].Value = 31;
                sheet.Cells["B4"].Value = 38;
                sheet.Cells["B5"].Value = 40;
                sheet.Cells["B8"].Formula = "FORECAST(30,A1:A5,B1:B5)";
                sheet.Calculate();
                var result = sheet.Cells["B8"].Value;

                Assert.AreEqual(10.607253, System.Math.Round((double)result, 6));
            }
        }

        [TestMethod]
        public void ForecastLinearTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 6;
                sheet.Cells["A2"].Value = 7;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 15;
                sheet.Cells["A5"].Value = 21;
                sheet.Cells["B1"].Value = 20;
                sheet.Cells["B2"].Value = 28;
                sheet.Cells["B3"].Value = 31;
                sheet.Cells["B4"].Value = 38;
                sheet.Cells["B5"].Value = 40;
                sheet.Cells["B8"].Formula = "FORECAST.LINEAR(30,A1:A5,B1:B5)";
                sheet.Calculate();
                var result = sheet.Cells["B8"].Value;

                Assert.AreEqual(10.607253, System.Math.Round((double)result, 6));
            }
        }

        [TestMethod]
        public void InterceptTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 1;
                sheet.Cells["A5"].Value = 8;
                sheet.Cells["B1"].Value = 6;
                sheet.Cells["B2"].Value = 5;
                sheet.Cells["B3"].Value = 11;
                sheet.Cells["B4"].Value = 7;
                sheet.Cells["B5"].Value = 5;
                sheet.Cells["B8"].Formula = "INTERCEPT(A1:A5,B1:B5)";
                sheet.Calculate();
                var result = sheet.Cells["B8"].Value;

                Assert.AreEqual(0.0483871, System.Math.Round((double)result, 7));
            }
        }
    }
}
