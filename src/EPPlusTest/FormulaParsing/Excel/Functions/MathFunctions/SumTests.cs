using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;

namespace EPPlusTest.FormulaParsing.Excel.Functions.MathFunctions
{
    [TestClass]
    public class SumTests
    {
        [TestMethod]
        public void ShouldTreatSingleBooleanValuesOrginatingFromEvaluationsAsNumbers()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "A";
                sheet.Cells["A2"].Value = "A";
                sheet.Cells["A4"].Formula = "SUM(A1=\"A\", A2=\"A\",A3=\"A\")";
                sheet.Calculate();
                var a4val = sheet.Cells["A4"].Value;
                Assert.AreEqual(2d, a4val);
            }
        }

        [TestMethod]
        public void ShouldTreatSingleBooleanValuesAsNumbers()
        {
            // the logic seems to be
            // that boolean values that originates from an evaluation with
            // cell addresses are not counted as numeric values.
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "TRUE";
                sheet.Cells["A2"].Formula = "TRUE";
                sheet.Cells["A4"].Formula = "SUM(A1,A2,A3)";
                sheet.Calculate();
                var a4val = sheet.Cells["A4"].Value;
                Assert.AreEqual(0d, a4val);
            }
        }

        [TestMethod]
        public void ShouldCountNumbers()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A4"].Formula = "SUM(A1,1)";
                sheet.Calculate();
                var a4val = sheet.Cells["A4"].Value;
                Assert.AreEqual(2d, a4val);
            }
        }

        [TestMethod]
        public void ShouldNotCountNumericStringsViaReference()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "1";
                sheet.Cells["A4"].Formula = "SUM(A1)";
                sheet.Calculate();
                var a4val = sheet.Cells["A4"].Value;
                Assert.AreEqual(0d, a4val);
            }
        }

        [TestMethod]
        public void ShouldCountNumericStringViaArgument()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A4"].Formula = "SUM(\"1\")";
                sheet.Calculate();
                var a4val = sheet.Cells["A4"].Value;
                Assert.AreEqual(1d, a4val);
            }
        }

        [TestMethod]
        public void ShouldCountDates()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = new DateTime(2023, 7, 7);
                sheet.Cells["A4"].Formula = "SUM(A1)";
                sheet.Calculate();
                var a4val = sheet.Cells["A4"].Value;
                Assert.AreEqual(45114d, a4val);
            }
        }

        [TestMethod]
        public void ShouldReturnErrorFromSingleCellArg()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "1/0";
                sheet.Cells["A4"].Formula = "SUM(A1)";
                sheet.Calculate();
                var a4val = sheet.Cells["A4"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), a4val);
            }
        }

        [TestMethod]
        public void ShouldReturnErrorFromMulticellRange()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Formula = "1/0";
                sheet.Cells["A4"].Formula = "SUM(A1:A2)";
                sheet.Calculate();
                var a4val = sheet.Cells["A4"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), a4val);
            }
        }


        [TestMethod]
        public void ShouldNotReturnErrorFromValidMulticellRange()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Value = -1;
                sheet.Cells["A4"].Formula = "SUM(A1:A2)";
                sheet.Calculate();
                var a4val = sheet.Cells["A4"].Value;
                Assert.AreEqual(1d,a4val);
            }
        }


        [TestMethod]
        public void SumPrecisionTest()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");

            ws.Cells["A2"].Value = 0.2;
            ws.Cells["B2"].Value = 21.9;
            ws.Cells["D2"].Value = 22.1;
            ws.Cells["E2"].Formula = "D2-SUM(A2:B2)";

            ws.Cells["A3"].Value = 0.2;
            ws.Cells["B3"].Value = -21.9;
            ws.Cells["E3"].Formula = "SUM(A3:B3)";

            ws.Cells["A4"].Value = -0.2;
            ws.Cells["B4"].Value = -21.9;
            ws.Cells["D4"].Value = -22.1;
            ws.Cells["E4"].Formula = "D4-SUM(A4:B4)";


            ws.Cells["A5"].Value = -0.2;
            ws.Cells["B5"].Value = 21.9;
            ws.Cells["D5"].Value = -22.1;
            ws.Cells["E5"].Formula = "D5-SUM(A5:B5)";

            p.Workbook.Calculate();

            Assert.AreEqual(0d, ws.Cells["E2"].Value);
            Assert.AreEqual(-21.7d, ws.Cells["E3"].Value);
            Assert.AreEqual(0d, ws.Cells["E4"].Value);
            Assert.AreEqual(-43.8, ws.Cells["E5"].Value);

            var r = RoundingHelper.GetSignificantFigures(-123456.0987654321d, 15);
            Assert.AreEqual(-123456.098765432, r);

            var r2 = RoundingHelper.GetSignificantFigures(0.000111111111111173d, 15);
            Assert.AreEqual(0.000111111111111173, r2);
        }
    }
}
