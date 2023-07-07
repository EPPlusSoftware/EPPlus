using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
    [TestClass]
    public class SumTests
    {
        // TODO: this needs a review and potentially redesign
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
        public void ShouldNotCountNumericStrings()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "1";
                sheet.Cells["A4"].Formula = "SUM(A1,\"1\")";
                sheet.Calculate();
                var a4val = sheet.Cells["A4"].Value;
                Assert.AreEqual(0d, a4val);
            }
        }

        [TestMethod]
        public void ShouldCountDates()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = new DateTime(2023, 7, 7);
                sheet.Cells["A4"].Formula = "SUM(A1,\"1\")";
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
                sheet.Cells["A1"].Formula = "1/0";
                sheet.Cells["A4"].Formula = "SUM(A1:A2)";
                sheet.Calculate();
                var a4val = sheet.Cells["A4"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), a4val);
            }
        }
    }
}
