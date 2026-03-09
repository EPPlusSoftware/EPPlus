using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class EntireColumnTests
    {
        private ExcelPackage _package;
        private ExcelWorkbook _workbook;

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _workbook = _package.Workbook;
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void FullColumnRefPlusScalar_EmptyColumn_ShouldReturnOneInFirstAndLastRow()
        {
            // Arrange
            var ws = _workbook.Worksheets.Add("Sheet1");
            // Column B is entirely empty/blank
            ws.Cells["C1"].Formula = "B:B+1";

            // Act
            ws.Calculate();

            // Assert - first row of spill result
            Assert.AreEqual(1d, ws.Cells["C1"].Value,
                "First row should be 0 + 1 = 1");
            var lastRow = ws.Dimension != null ? ws.Dimension.End.Row : 1;
            Assert.AreEqual(ExcelPackage.MaxRows, ws.Dimension.End.Row,
                "Spill should extend to the last row of the worksheet");
            Assert.AreEqual(1d, ws.Cells["C" + lastRow].Value,
                "Last row of spill should be 0 + 1 = 1");
        }

        [TestMethod]
        public void FullColumnRefPlusScalar_WithData_PhysicalRowsCalculatedAndVirtualRowsGetDefault()
        {
            // Arrange - B has data in rows 1-3, formula in C1
            var ws = _workbook.Worksheets.Add("Sheet1");
            ws.Cells["B1"].Value = 10d;
            ws.Cells["B2"].Value = 20d;
            ws.Cells["B3"].Value = 30d;
            ws.Cells["C1"].Formula = "B:B+5";

            // Act
            ws.Calculate();

            // Assert - physical rows get their actual calculated values
            Assert.AreEqual(15d, ws.Cells["C1"].Value);
            Assert.AreEqual(25d, ws.Cells["C2"].Value);
            Assert.AreEqual(35d, ws.Cells["C3"].Value);
            // Virtual rows beyond data: empty + 5 = 5
            Assert.AreEqual(5d, ws.Cells["C4"].Value,
                "First virtual row should be 0 + 5 = 5");
            Assert.AreEqual(5d, ws.Cells["C" + ExcelPackage.MaxRows].Value,
                "Last row should be 0 + 5 = 5");
        }

        [TestMethod]
        public void FullColumnRefEqualsString_VirtualRowsShouldBeFalse()
        {
            // Arrange - comparison operator: A:A="Hello"
            // Virtual rows are empty, empty != "Hello" => FALSE
            var ws = _workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = "Hello";
            ws.Cells["A2"].Value = "World";
            ws.Cells["A3"].Value = "Hello";
            ws.Cells["B1"].Formula = "A:A=\"Hello\"";

            // Act
            ws.Calculate();

            // Assert - physical rows
            Assert.AreEqual(true, ws.Cells["B1"].Value);
            Assert.AreEqual(false, ws.Cells["B2"].Value);
            Assert.AreEqual(true, ws.Cells["B3"].Value);
            // Virtual rows: empty = "Hello" => FALSE
            Assert.AreEqual(false, ws.Cells["B4"].Value,
                "Virtual row: empty = \"Hello\" should be FALSE");
            Assert.AreEqual(false, ws.Cells["B" + ExcelPackage.MaxRows].Value,
                "Last row should also be FALSE");
        }

        [TestMethod]
        public void TwoFullColumnRefsMultiplied_VirtualRowsShouldBeZero()
        {
            // Arrange - (A:A=x) * (B:B=y) pattern used in MATCH criteria
            var ws = _workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = "Alpha";
            ws.Cells["A2"].Value = "Beta";
            ws.Cells["B1"].Value = "X";
            ws.Cells["B2"].Value = "Y";
            ws.Cells["C1"].Value = 100d;
            ws.Cells["C2"].Value = 200d;

            // MATCH(1, (A:A="Alpha")*(B:B="X"), 0) should find row 1
            ws.Cells["E1"].Formula = "MATCH(1,(A:A=\"Alpha\")*(B:B=\"X\"),0)";
            // INDEX to retrieve the value
            ws.Cells["F1"].Formula = "INDEX(C:C,MATCH(1,(A:A=\"Alpha\")*(B:B=\"X\"),0))";

            // Act
            ws.Calculate();

            // Assert
            Assert.AreEqual(1, ws.Cells["E1"].Value,
                "MATCH should find row 1");
            Assert.AreEqual(100d, ws.Cells["F1"].Value,
                "INDEX should return 100 for Alpha+X");
        }

        [TestMethod]
        public void CrossSheetFullColumnRef_IndexMatch_ShouldWork()
        {
            // Arrange - the original problem pattern from the design doc
            var dataWs = _workbook.Worksheets.Add("Data");
            dataWs.Cells["A1"].Value = "Item1";
            dataWs.Cells["A2"].Value = "Item2";
            dataWs.Cells["A3"].Value = "Item1";
            dataWs.Cells["B1"].Value = "Day Shift";
            dataWs.Cells["B2"].Value = "Day Shift";
            dataWs.Cells["B3"].Value = "Night Shift";
            dataWs.Cells["C1"].Value = 50d;
            dataWs.Cells["C2"].Value = 75d;
            dataWs.Cells["C3"].Value = 90d;

            var ws = _workbook.Worksheets.Add("Formulas");
            ws.Cells["A1"].Value = "Item1";
            ws.Cells["B1"].Formula =
                "IFERROR(INDEX(Data!$C:$C,MATCH(1,(Data!$A:$A=A1)*(Data!$B:$B=\"Day Shift\"),0)),\"-\")";
            // Also test a lookup that should NOT match
            ws.Cells["A2"].Value = "NoMatch";
            ws.Cells["B2"].Formula =
                "IFERROR(INDEX(Data!$C:$C,MATCH(1,(Data!$A:$A=A2)*(Data!$B:$B=\"Day Shift\"),0)),\"-\")";

            // Act
            _workbook.Calculate();

            // Assert
            Assert.AreEqual(50d, ws.Cells["B1"].Value,
                "Should find Item1 + Day Shift => 50");
            Assert.AreEqual("-", ws.Cells["B2"].Value,
                "NoMatch should fall through to IFERROR => \"-\"");
        }
        [TestMethod]
        public void ScalarDivideFullColumnRef_VirtualRowsShouldBeDivByZeroError()
        {
            // Arrange - 1/B:B where B has data in rows 1-2
            // Virtual default: 1 / null = 1 / 0 = #DIV/0!
            var ws = _workbook.Worksheets.Add("Sheet1");
            ws.Cells["B1"].Value = 2d;
            ws.Cells["B2"].Value = 4d;
            ws.Cells["C1"].Formula = "1/B:B";

            // Act
            ws.Calculate();

            // Assert - physical rows
            Assert.AreEqual(0.5d, ws.Cells["C1"].Value);
            Assert.AreEqual(0.25d, ws.Cells["C2"].Value);
            // Virtual rows: 1 / 0 => #DIV/0!
            var virtualVal = ws.Cells["C3"].Value;
            Assert.IsInstanceOfType(virtualVal, typeof(ExcelErrorValue),
                "Virtual row: 1/0 should produce an error");
            Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)virtualVal).Type,
                "Error should be #DIV/0!");
            // Last row should also be #DIV/0!
            var lastVal = ws.Cells["C" + ExcelPackage.MaxRows].Value;
            Assert.IsInstanceOfType(lastVal, typeof(ExcelErrorValue),
                "Last row should also be #DIV/0!");
        }

        [TestMethod]
        public void FullColumnRefConcat_VirtualRowsShouldConcatEmpty()
        {
            // Arrange - A:A&"!" via the Concat operator
            // Virtual default: "" & "!" = "!"
            var ws = _workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = "Hello";
            ws.Cells["A2"].Value = "World";
            ws.Cells["B1"].Formula = "A:A&\"!\"";

            // Act
            ws.Calculate();

            // Assert - physical rows
            Assert.AreEqual("Hello!", ws.Cells["B1"].Value);
            Assert.AreEqual("World!", ws.Cells["B2"].Value);
            // Virtual rows: "" & "!" = "!"
            Assert.AreEqual("!", ws.Cells["B3"].Value,
                "Virtual row: empty & \"!\" should be \"!\"");
            Assert.AreEqual("!", ws.Cells["B" + ExcelPackage.MaxRows].Value,
                "Last row should also be \"!\"");
        }

        [TestMethod]
        public void NegateFullColumnRef_VirtualRowsShouldBeZero()
        {
            // Arrange - negation: -A:A
            // Virtual default: -(null) = -(0) = 0
            var ws = _workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = 5d;
            ws.Cells["A2"].Value = -3d;
            ws.Cells["B1"].Formula = "-A:A";

            // Act
            ws.Calculate();

            // Assert - physical rows
            Assert.AreEqual(-5d, ws.Cells["B1"].Value);
            Assert.AreEqual(3d, ws.Cells["B2"].Value);
            // Virtual rows: -(0) = 0
            Assert.AreEqual(0d, ws.Cells["B3"].Value,
                "Virtual row: -0 should be 0");
            Assert.AreEqual(0d, ws.Cells["B" + ExcelPackage.MaxRows].Value,
                "Last row should also be 0");
        }
    }
}