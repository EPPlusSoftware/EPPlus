using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class NegatorTests
    {
        [TestMethod]
        public void NegateNamedRange()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
                sheet1.Cells["A1"].Value = 123456;

                sheet1.Names.Add("MyRange", sheet1.Cells["A1"]);

                sheet1.Cells["C3"].Formula = "-MyRange";

                pck.Workbook.Calculate();

                Assert.AreEqual(-123456, sheet1.Cells["C3"].GetValue<double>(), 1E-5); //ERROR: evaluates to 123456
            }
        }
        [TestMethod]
        public void NegateNamedRangePlusNamedRange()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
                sheet1.Cells["A1"].Value = 123456;
                sheet1.Cells["B1"].Value = 3;

                sheet1.Names.Add("MyRange", sheet1.Cells["A1"]);
                sheet1.Names.Add("Another", sheet1.Cells["B1"]);

                sheet1.Cells["C3"].Formula = "-MyRange+Another";

                pck.Workbook.Calculate();

                Assert.AreEqual(-123453, sheet1.Cells["C3"].GetValue<double>(), 1E-5); //ERROR: evaluates to 123459
            }
        }

        [TestMethod]
        public void NegateNamedRangePlusNamedRange_WithParenthesis()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
                sheet1.Cells["A1"].Value = 123456;
                sheet1.Cells["B1"].Value = 3;

                sheet1.Names.Add("MyRange", sheet1.Cells["A1"]);
                sheet1.Names.Add("Another", sheet1.Cells["B1"]);

                sheet1.Cells["C3"].Formula = "-(MyRange+Another)";

                pck.Workbook.Calculate();

                Assert.AreEqual(-123459, sheet1.Cells["C3"].GetValue<double>(), 1E-5); //ERROR: evaluates to 123459
            }
        }

        [TestMethod]
        public void DoubleNegateNamedRangePlusNamedRange_WithParenthesis()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
                sheet1.Cells["A1"].Value = 123456;
                sheet1.Cells["B1"].Value = 3;

                sheet1.Names.Add("MyRange", sheet1.Cells["A1"]);
                sheet1.Names.Add("Another", sheet1.Cells["B1"]);

                sheet1.Cells["C3"].Formula = "--(MyRange+Another)";

                pck.Workbook.Calculate();

                Assert.AreEqual(123459, sheet1.Cells["C3"].GetValue<double>(), 1E-5); //ERROR: evaluates to 123459
            }
        }

        [TestMethod]
        public void NegateMultiCellNamedRange()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
                sheet1.Cells["A1"].Value = 1;
                sheet1.Cells["B1"].Value = 2;
                sheet1.Cells["A2"].Value = 3;
                sheet1.Cells["B2"].Value = "abc";

                sheet1.Names.Add("MyRange", sheet1.Cells["A1:B2"]);

                sheet1.Cells["E4"].Formula = "-MyRange";

                pck.Workbook.Calculate();

                Assert.AreEqual(-1, sheet1.Cells["E4"].GetValue<double>(), 1E-5);
                Assert.AreEqual(-2, sheet1.Cells["F4"].GetValue<double>(), 1E-5);
                Assert.AreEqual(-3, sheet1.Cells["E5"].GetValue<double>(), 1E-5);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet1.Cells["F5"].Value);
            }
        }

        [TestMethod]

        public void NegateMultiCellRange()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
                sheet1.Cells["A1"].Value = 1;
                sheet1.Cells["B1"].Value = 2;
                sheet1.Cells["A2"].Value = 3;
                sheet1.Cells["B2"].Value = 4;

                sheet1.Cells["E4"].Formula = "-A1:B2";

                pck.Workbook.Calculate();

                Assert.AreEqual(-1, sheet1.Cells["E4"].GetValue<double>(), 1E-5);
                Assert.AreEqual(-2, sheet1.Cells["F4"].GetValue<double>(), 1E-5);
                Assert.AreEqual(-3, sheet1.Cells["E5"].GetValue<double>(), 1E-5);
                Assert.AreEqual(-4, sheet1.Cells["F5"].GetValue<double>(), 1E-5);
            }
        }

        [TestMethod]
        public void DoubleNegateMultiCellRange()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
                sheet1.Cells["A1"].Value = 1;
                sheet1.Cells["B1"].Value = 2;
                sheet1.Cells["A2"].Value = 3;
                sheet1.Cells["B2"].Value = 4;

                sheet1.Cells["E4"].Formula = "--A1:B2";

                pck.Workbook.Calculate();

                Assert.AreEqual(1, sheet1.Cells["E4"].GetValue<double>(), 1E-5);
                Assert.AreEqual(2, sheet1.Cells["F4"].GetValue<double>(), 1E-5);
                Assert.AreEqual(3, sheet1.Cells["E5"].GetValue<double>(), 1E-5);
                Assert.AreEqual(4, sheet1.Cells["F5"].GetValue<double>(), 1E-5);
            }
        }

        [TestMethod]
        public void NegateMultiCellRange_WithString()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
                sheet1.Cells["A1"].Value = 1;
                sheet1.Cells["B1"].Value = 2;
                sheet1.Cells["A2"].Value = 3;
                sheet1.Cells["B2"].Value = "abc";

                sheet1.Cells["E4"].Formula = "-A1:B2";

                pck.Workbook.Calculate();

                Assert.AreEqual(-1, sheet1.Cells["E4"].GetValue<double>(), 1E-5);
                Assert.AreEqual(-2, sheet1.Cells["F4"].GetValue<double>(), 1E-5);
                Assert.AreEqual(-3, sheet1.Cells["E5"].GetValue<double>(), 1E-5);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet1.Cells["F5"].Value);
            }
        }
        [TestMethod]
        public void DoubleNegationsWithCells()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("sheet1");
                ws.Cells["A1"].Value = -1.5;
                ws.Cells["B1"].Value = -5;
                ws.Cells["C1"].Value = 1.5;
                ws.Cells["D1"].Formula = "IF((A1+B1)<0,(-A1+-B1)*C1,0)";
                ws.Cells["E1"].Formula = "IF((A1+B1)<0,(-A1+--B1)*C1,0)";
                ws.Cells["F1"].Formula = "IF((A1+B1)<0,(-A1+-(-B1))*C1,0)";
                ws.Cells["G1"].Formula = "IF((A1+B1)<0,(--A1+-(-B1))*C1,0)";
                ws.Calculate();

                Assert.AreEqual(9.75, ws.Cells["D1"].Value);
                Assert.AreEqual(-5.25, ws.Cells["E1"].Value);
                Assert.AreEqual(-5.25, ws.Cells["F1"].Value);
                Assert.AreEqual(-9.75, ws.Cells["G1"].Value);
            }
        }
		[TestMethod]
		public void NegateAnEmptyCellShouldReturnZero()
		{
			using (var p = new ExcelPackage())
			{
				var ws = p.Workbook.Worksheets.Add("Sheet1");

				ws.Cells["A1"].Formula = "-B1";
				ws.Calculate();

				Assert.AreEqual(0d, ws.Cells["A1"].Value);
			}
		}

	}
}
