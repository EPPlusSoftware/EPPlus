using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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


            ws.Cells["A3"].Value = 1;
            ws.Cells["B3"].Value = 9000;
            ws.Cells["C3"].Formula = "A3/B3";
            ws.Cells["D3"].Formula = "1+C3";
            ws.Cells["E3"].Formula = "D3-1";

            p.Workbook.Calculate(
                                new ExcelCalculationOption
                                {
                                    PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel
                                });

            var res = float.Parse( ws.Cells["D2"].Value.ToString()) - (float.Parse( ws.Cells["A2"].Value.ToString()) + float.Parse(ws.Cells["B2"].Value.ToString()));

            Assert.AreEqual(0d, ws.Cells["E2"].Value);
            //Assert.AreEqual(0, ws.Cells["E3"].Value);
        }
    }
}
