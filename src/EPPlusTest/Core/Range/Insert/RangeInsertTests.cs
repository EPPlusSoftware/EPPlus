using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EPPlusTest.Core.Range.Insert
{
    [TestClass]
    public class RangeInsertTests : TestBase
    {
        public static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("WorksheetRangeInsert.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void ValidateFormulasAfterInsertRow()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("InsertRow_Sheet1");
            var ws2 = _pck.Workbook.Worksheets.Add("InsertRow_Sheet2");
            ws.Cells["A1"].Formula = "Sum(C5:C10)";
            ws.Cells["B1:B2"].Formula = "Sum(C5:C10)";
            ws2.Cells["A1"].Formula = "Sum(InsertRow_Sheet1!C5:C10)";
            ws2.Cells["B1:B2"].Formula = "Sum(InsertRow_Sheet1!C5:C10)";

            //Act
            ws.InsertRow(3, 1);

            //Assert
            Assert.AreEqual(1, ws._sharedFormulas.Count);
            Assert.AreEqual(1, ws._sharedFormulas.First().Key);
            Assert.AreEqual("Sum(C6:C11)", ws.Cells["A1"].Formula);
            Assert.AreEqual("Sum(C6:C11)", ws.Cells["B1"].Formula);
            Assert.AreEqual("Sum(C7:C12)", ws.Cells["B2"].Formula);

            Assert.AreEqual("Sum(InsertRow_Sheet1!C6:C11)", ws2.Cells["A1"].Formula);
            Assert.AreEqual("Sum(InsertRow_Sheet1!C6:C11)", ws2.Cells["B1"].Formula);
            Assert.AreEqual("Sum(InsertRow_Sheet1!C7:C12)", ws2.Cells["B2"].Formula);
        }
        [TestMethod]
        public void ValidateFormulasAfterInsert2Rows()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("Insert2Rows_Sheet1");
            var ws2 = _pck.Workbook.Worksheets.Add("Insert2Rows_Sheet2");
            ws.Cells["A1"].Formula = "Sum(C5:C10)";
            ws.Cells["B1:B2"].Formula = "Sum(C5:C10)";
            ws2.Cells["A1"].Formula = "Sum(Insert2Rows_Sheet1!C5:C10)";
            ws2.Cells["B1:B2"].Formula = "Sum(Insert2Rows_Sheet1!C5:C10)";

            //Act
            ws.InsertRow(3, 2);

            //Assert
            Assert.AreEqual("Sum(C7:C12)", ws.Cells["A1"].Formula);
            Assert.AreEqual("Sum(C7:C12)", ws.Cells["B1"].Formula);
            Assert.AreEqual("Sum(C8:C13)", ws.Cells["B2"].Formula);

            Assert.AreEqual("Sum(Insert2Rows_Sheet1!C7:C12)", ws2.Cells["A1"].Formula);
            Assert.AreEqual("Sum(Insert2Rows_Sheet1!C7:C12)", ws2.Cells["B1"].Formula);
            Assert.AreEqual("Sum(Insert2Rows_Sheet1!C8:C13)", ws2.Cells["B2"].Formula);
        }
        [TestMethod]
        public void ValidateFormulasAfterInsertColumn()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("InsertColumn_Sheet1");
            var ws2 = _pck.Workbook.Worksheets.Add("InsertColumn_Sheet2");
            ws.Cells["A1"].Formula = "Sum(E1:J1)";
            ws.Cells["B1:C1"].Formula = "Sum(E1:J1)";
            ws2.Cells["A1"].Formula = "Sum(InsertColumn_Sheet1!E1:J1)";
            ws2.Cells["B1:C1"].Formula = "Sum(InsertColumn_Sheet1!E1:J1)";

            //Act
            ws.InsertColumn(4, 1);

            //Assert
            Assert.AreEqual("Sum(F1:K1)", ws.Cells["A1"].Formula);
            Assert.AreEqual("Sum(F1:K1)", ws.Cells["B1"].Formula);
            Assert.AreEqual("Sum(G1:L1)", ws.Cells["C1"].Formula);

            Assert.AreEqual("Sum(InsertColumn_Sheet1!F1:K1)", ws2.Cells["A1"].Formula);
            Assert.AreEqual("Sum(InsertColumn_Sheet1!F1:K1)", ws2.Cells["B1"].Formula);
            Assert.AreEqual("Sum(InsertColumn_Sheet1!G1:L1)", ws2.Cells["C1"].Formula);
        }
        [TestMethod]
        public void ValidateFormulasAfterInsert2Columns()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("Insert2Columns_Sheet1");
            var ws2 = _pck.Workbook.Worksheets.Add("Insert2Columns_Sheet2");
            ws.Cells["A1"].Formula = "Sum(E1:J1)";
            ws.Cells["B1:C1"].Formula = "Sum(E1:J1)";
            ws2.Cells["A1"].Formula = "Sum(Insert2Columns_Sheet1!E1:J1)";
            ws2.Cells["B1:C1"].Formula = "Sum(Insert2Columns_Sheet1!E1:J1)";

            //Act
            ws.InsertColumn(4, 2);

            //Assert
            Assert.AreEqual("Sum(G1:L1)", ws.Cells["A1"].Formula);
            Assert.AreEqual("Sum(G1:L1)", ws.Cells["B1"].Formula);
            Assert.AreEqual("Sum(H1:M1)", ws.Cells["C1"].Formula);

            Assert.AreEqual("Sum(Insert2Columns_Sheet1!G1:L1)", ws2.Cells["A1"].Formula);
            Assert.AreEqual("Sum(Insert2Columns_Sheet1!G1:L1)", ws2.Cells["B1"].Formula);
            Assert.AreEqual("Sum(Insert2Columns_Sheet1!H1:M1)", ws2.Cells["C1"].Formula);
        }
        [TestMethod]
        public void InsertingColumnIntoTable()
        {
            using (var p = new ExcelPackage())
            {
                //Setup
                var ws = p.Workbook.Worksheets.Add("InsertColumnTable");
                LoadTestdata(ws);
                var tbl = ws.Tables.Add(ws.Cells[1, 1, 100, 5], "Table1");
                //Act
                ws.InsertColumn(2, 1);

                //Assert
                Assert.AreEqual(6, tbl.Columns.Count);
                Assert.AreEqual("Date", tbl.Columns[0].Name);
                Assert.AreEqual("Column2", tbl.Columns[1].Name);
                Assert.AreEqual("NumValue", tbl.Columns[2].Name);
                Assert.AreEqual("StrValue", tbl.Columns[3].Name);
                Assert.AreEqual("NumFormattedValue", tbl.Columns[4].Name);
                Assert.AreEqual("Column5", tbl.Columns[5].Name);
            }
        }
        [TestMethod]
        public void InsertingRowIntoTable()
        {
            using (var p = new ExcelPackage())
            {
                //Setup
                var ws = p.Workbook.Worksheets.Add("InsertRowTable");
                LoadTestdata(ws);
                var tbl = ws.Tables.Add(ws.Cells[1, 1, 100, 5], "Table1");
                //Act
                ws.InsertRow(1, 1);
                ws.InsertRow(3, 1);
                ws.InsertRow(103, 1);

                //Assert
                Assert.AreEqual("A2:E102", tbl.Address.Address);
            }
        }
        [TestMethod]
        public void ValidateValuesAfterInsertRowInRangeShiftDown()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("InsertRangeDown");
            ws.Cells["A1"].Value = "A1";
            ws.Cells["B1"].Value = "B1";
            ws.Cells["C1"].Value = "C1";

            //Act
            ws.Cells["B1"].Insert(eShiftTypeInsert.Down);

            //Assert
            Assert.AreEqual("A1", ws.Cells["A1"].Value);
            Assert.IsNull(ws.Cells["B1"].Value);
            Assert.AreEqual("B1", ws.Cells["B2"].Value);
            Assert.AreEqual("C1", ws.Cells["C1"].Value);
        }
        [TestMethod]
        public void ValidateValuesAfterInsertRowInRangeShiftDownTwoRows()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("InsertRangeDownTwoRows");
            ws.Cells["A1"].Value = "A1";
            ws.Cells["B1"].Value = "B1";
            ws.Cells["C1"].Value = "C1";
            ws.Cells["D1"].Value = "D1";
            ws.Cells["A2"].Value = "A2";
            ws.Cells["B2"].Value = "B2";
            ws.Cells["C2"].Value = "C2";
            ws.Cells["D2"].Value = "D2";

            //Act
            ws.Cells["B1:C2"].Insert(eShiftTypeInsert.Down);

            //Assert
            Assert.AreEqual("A1", ws.Cells["A1"].Value);
            Assert.IsNull(ws.Cells["B1"].Value);
            Assert.IsNull(ws.Cells["C1"].Value);
            Assert.IsNull(ws.Cells["B2"].Value);
            Assert.IsNull(ws.Cells["C2"].Value);
            Assert.AreEqual("B1", ws.Cells["B3"].Value);
            Assert.AreEqual("C1", ws.Cells["C3"].Value);
            Assert.AreEqual("A2", ws.Cells["A2"].Value);
            Assert.AreEqual("D2", ws.Cells["D2"].Value);
        }
        [TestMethod]
        public void ValidateValuesAfterInsertRowInRangeShiftRight()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("InsertRangeRight");
            ws.Cells["A1"].Value = "A1";
            ws.Cells["B1"].Value = "B1";
            ws.Cells["C1"].Value = "C1";

            //Act
            ws.Cells["B1"].Insert(eShiftTypeInsert.Right);

            //Assert
            Assert.AreEqual("A1", ws.Cells["A1"].Value);
            Assert.IsNull(ws.Cells["B1"].Value);
            Assert.AreEqual("B1", ws.Cells["C1"].Value);
            Assert.AreEqual("C1", ws.Cells["D1"].Value);
        }
        [TestMethod]
        public void ValidateValuesAfterInsertRowInRangeShiftRightTwoRows()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("InsertRangeRightTwoRows");
            ws.Cells["A1"].Value = "A1";
            ws.Cells["B1"].Value = "B1";
            ws.Cells["C1"].Value = "C1";
            ws.Cells["D1"].Value = "D1";
            ws.Cells["A2"].Value = "A2";
            ws.Cells["B2"].Value = "B2";
            ws.Cells["C2"].Value = "C2";
            ws.Cells["D2"].Value = "D2";

            //Act
            ws.Cells["B1:C2"].Insert(eShiftTypeInsert.Right);

            //Assert
            Assert.AreEqual("A1", ws.Cells["A1"].Value);
            Assert.IsNull(ws.Cells["B1"].Value);
            Assert.IsNull(ws.Cells["C1"].Value);
            Assert.IsNull(ws.Cells["B2"].Value);
            Assert.IsNull(ws.Cells["C2"].Value);
            Assert.AreEqual("B1", ws.Cells["D1"].Value);
            Assert.AreEqual("C1", ws.Cells["E1"].Value);
            Assert.AreEqual("B2", ws.Cells["D2"].Value);
            Assert.AreEqual("C2", ws.Cells["E2"].Value);
            Assert.AreEqual("D2", ws.Cells["F2"].Value);
        }
        [TestMethod]
        public void ValidateCommentsAfterInsertShiftDown()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("InsertRangeCommentsDown");
            ws.Cells["A1"].AddComment("Comment A1", "EPPlus");
            ws.Cells["B1"].AddComment("Comment B1", "EPPlus");
            ws.Cells["C1"].AddComment("Comment C1", "EPPlus");

            //Act
            ws.Cells["A1"].Insert(eShiftTypeInsert.Down);

            //Assert
            Assert.IsNull(ws.Cells["A1"].Comment);
            Assert.AreEqual("Comment A1", ws.Cells["A2"].Comment.Text);
            Assert.AreEqual("Comment B1", ws.Cells["B1"].Comment.Text);
            Assert.AreEqual("Comment C1", ws.Cells["C1"].Comment.Text);
        }
        [TestMethod]
        public void ValidateCommentsAfterInsertShiftRight()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("InsertRangeCommentsRight");
            ws.Cells["A1"].AddComment("Comment A1", "EPPlus");  
            ws.Cells["B1"].AddComment("Comment B1", "EPPlus");
            ws.Cells["C1"].AddComment("Comment C1", "EPPlus");

            //Act
            ws.Cells["A1"].Insert(eShiftTypeInsert.Right);

            //Assert
            Assert.IsNull(ws.Cells["A1"].Comment);
            Assert.AreEqual("Comment A1", ws.Cells["B1"].Comment.Text);
            Assert.AreEqual("Comment B1", ws.Cells["C1"].Comment.Text);
            Assert.AreEqual("Comment C1", ws.Cells["D1"].Comment.Text);
            Assert.IsNull(ws.Cells["A2"].Comment);
        }
        [TestMethod]
        public void ValidateNameAfterInsertShiftDown()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("InsertRangeNamesDown");
            ws.Names.Add("NameA1", ws.Cells["A1"]);
            ws.Names.Add("NameB1", ws.Cells["B1"]);
            ws.Names.Add("NameC1", ws.Cells["C1"]);

            //Act
            ws.Cells["A1"].Insert(eShiftTypeInsert.Down);

            //Assert
            Assert.AreEqual("$A$2", ws.Names["NameA1"].Address);
            Assert.AreEqual("$B$1", ws.Names["NameB1"].Address);
            Assert.AreEqual("$C$1", ws.Names["NameC1"].Address);
        }
        [TestMethod]
        public void ValidateNameAfterInsertShiftDown_MustBeInsideRange()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("InsertRangeInsideNamesDown");
            ws.Names.Add("NameA2B4", ws.Cells["A2:B4"]);
            ws.Names.Add("NameB2D3", ws.Cells["B2:D3"]);
            ws.Names.Add("NameC1F3", ws.Cells["C1:F3"]);

            //Act
            ws.Cells["A2:B3"].Insert(eShiftTypeInsert.Down);

            //Assert
            Assert.AreEqual("$A$4:$B$6", ws.Names["NameA2B4"].Address);
            Assert.AreEqual("$B$2:$D$3", ws.Names["NameB2D3"].Address);
            Assert.AreEqual("$C$1:$F$3", ws.Names["NameC1F3"].Address);

            ws.Cells["B2:D5"].Insert(eShiftTypeInsert.Down);
            Assert.AreEqual("$A$4:$B$6", ws.Names["NameA2B4"].Address);
            Assert.AreEqual("$B$6:$D$7", ws.Names["NameB2D3"].Address);
            Assert.AreEqual("$C$1:$F$3", ws.Names["NameC1F3"].Address);

            ws.Cells["B2:F2"].Insert(eShiftTypeInsert.Down);
            Assert.AreEqual("$A$4:$B$6", ws.Names["NameA2B4"].Address);
            Assert.AreEqual("$B$7:$D$8", ws.Names["NameB2D3"].Address);
            Assert.AreEqual("$C$1:$F$4", ws.Names["NameC1F3"].Address);
        }

        [TestMethod]
        public void ValidateNamesAfterInsertShiftRight_MustBeInsideRange()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("InsertRangeInsideNamesRight");
            ws.Names.Add("NameB1D2", ws.Cells["B1:D2"]);
            ws.Names.Add("NameB2C4", ws.Cells["B2:D4"]);
            ws.Names.Add("NameA3C6", ws.Cells["A3:C6"]);

            //Act
            ws.Cells["B1:C2"].Insert(eShiftTypeInsert.Right);

            //Assert
            Assert.AreEqual("$D$1:$F$2", ws.Names["NameB1D2"].Address);
            Assert.AreEqual("$B$2:$D$4", ws.Names["NameB2C4"].Address);
            Assert.AreEqual("$A$3:$C$6", ws.Names["NameA3C6"].Address);

            ws.Cells["B2:D5"].Insert(eShiftTypeInsert.Down);
            Assert.AreEqual("$D$1:$F$2", ws.Names["NameB1D2"].Address);
            Assert.AreEqual("$B$6:$D$8", ws.Names["NameB2C4"].Address);
            Assert.AreEqual("$A$3:$C$6", ws.Names["NameA3C6"].Address);
        }

        [TestMethod]
        public void ValidateSharedFormulasInsertShiftDown()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("InsertRangeFormulaDown");
            ws.Cells["B1:D2"].Formula = "A1";
            ws.Cells["C3:F4"].Formula = "A1";

            //Act
            ws.Cells["B1"].Insert(eShiftTypeInsert.Down);

            //Assert
            Assert.AreEqual("A1", ws.Cells["B2"].Formula);
            Assert.AreEqual("A2", ws.Cells["B3"].Formula);
            Assert.AreEqual("A1", ws.Cells["C3"].Formula);
            Assert.AreEqual("A1", ws.Cells["C3"].Formula);
            Assert.AreEqual("B2", ws.Cells["D3"].Formula);
            Assert.AreEqual("C1", ws.Cells["E3"].Formula);
            Assert.AreEqual("D1", ws.Cells["F3"].Formula);


            Assert.AreEqual("A1", ws.Cells["C3"].Formula);
            Assert.AreEqual("D2", ws.Cells["F4"].Formula);

        }
        [TestMethod]
        public void ValidateSharedFormulasInsertShiftRight()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("InsertRangeFormulaRight");
            ws.Cells["B1:D2"].Formula = "A1";
            ws.Cells["C3:F4"].Formula = "A1";

            //Act
            ws.Cells["B1"].Insert(eShiftTypeInsert.Right);

            //Assert
            Assert.AreEqual("", ws.Cells["B1"].Formula);
            Assert.AreEqual("A1", ws.Cells["C1"].Formula);
            Assert.AreEqual("C1", ws.Cells["D1"].Formula);
            Assert.AreEqual("D1", ws.Cells["E1"].Formula);
            Assert.AreEqual("A2", ws.Cells["B2"].Formula);
            Assert.AreEqual("A1", ws.Cells["C3"].Formula);
            Assert.AreEqual("A1", ws.Cells["C3"].Formula);
            Assert.AreEqual("C1", ws.Cells["D3"].Formula);
            Assert.AreEqual("D1", ws.Cells["E3"].Formula);


            Assert.AreEqual("A1", ws.Cells["C3"].Formula);
            Assert.AreEqual("D2", ws.Cells["F4"].Formula);

        }
        [TestMethod]
        public void ValidateInsertMergedCellsDown()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("MergedCells");
                ws.Cells["C3:E4"].Merge = true;
                ws.Cells["C2:E2"].Insert(eShiftTypeInsert.Down);

                Assert.AreEqual("C4:E5", ws.MergedCells[0]);
            }
        }
        [TestMethod]
        public void ValidateInsertMergedCellsRight()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("MergedCells");
                ws.Cells["C2:E3"].Merge = true;
                ws.Cells["B2:B3"].Insert(eShiftTypeInsert.Right);

                Assert.AreEqual("D2:F3", ws.MergedCells[0]);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ValidateInsertIntoMergedCellsPartialRightThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("MergedCells");
                ws.Cells["B2:D3"].Merge = true;
                ws.Cells["A2"].Insert(eShiftTypeInsert.Right);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ValidateInsertIntoMergedCellsPartialDownThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("MergedCells");
                ws.Cells["B2:D3"].Merge = true;
                ws.Cells["C1"].Insert(eShiftTypeInsert.Down);
            }
        }
        [TestMethod]
        public void ValidateDeleteEntireMergeCells()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("MergedCells");
                ws.Cells["B2:D3"].Merge = true;
                Assert.AreEqual(1, ws.MergedCells.Count);
                Assert.AreEqual("B2:D3", ws.MergedCells[0]);
                Assert.IsTrue(ws.Cells["B2"].Merge);
                Assert.IsTrue(ws.Cells["D3"].Merge);
                ws.Cells["B2:D3"].Delete(eShiftTypeDelete.Up);
                Assert.AreEqual(1, ws.MergedCells.Count);
                Assert.IsFalse(ws.Cells["B2"].Merge);
                Assert.IsFalse(ws.Cells["D3"].Merge);
                Assert.IsNull(ws.MergedCells[0]);
            }
        }
        [TestMethod]
        public void ValidateInsertMergedCellsShouldBeShifted()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("MergedCells");
                ws.Cells["B3:D3"].Merge = true;
                ws.Cells["B3:D3"].Insert(eShiftTypeInsert.Down);

                Assert.AreEqual("B4:D4", ws.MergedCells[0]);
                Assert.IsFalse(ws.Cells["B3"].Merge);
                Assert.IsFalse(ws.Cells["C3"].Merge);
                Assert.IsFalse(ws.Cells["D3"].Merge);

                Assert.IsTrue(ws.Cells["B4"].Merge);
                Assert.IsTrue(ws.Cells["C4"].Merge);
                Assert.IsTrue(ws.Cells["D4"].Merge);

                ws.InsertRow(3, 1);
                Assert.AreEqual("B5:D5", ws.MergedCells[0]);
            }
        }
        [TestMethod]
        public void ValidateInsertIntoMergedCellsPartialRightShouldNotThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("MergedCells");
                ws.Cells["B2:D3"].Merge = true;
                ws.Cells["C1"].Insert(eShiftTypeInsert.Right);
            }
        }
        [TestMethod]
        public void ValidateInsertIntoMergedCellsPartialDownShouldNotThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("MergedCells");
                ws.Cells["B2:D3"].Merge = true;
                ws.Cells["A2"].Insert(eShiftTypeInsert.Down);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ValidateInsertIntoTablePartialRightThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("TableDelete");
                ws.Tables.Add(ws.Cells["B2:D3"], "table1");
                ws.Cells["A2"].Insert(eShiftTypeInsert.Right);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ValidateInsertIntoTablePartialDownThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("TableDelete");
                ws.Tables.Add(ws.Cells["B2:D3"], "table1");
                ws.Cells["C1"].Insert(eShiftTypeInsert.Down);
            }
        }
        [TestMethod]
        public void ValidateInsertIntoTablePartialRightShouldNotThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("TableDelete");
                ws.Tables.Add(ws.Cells["B2:D3"], "table1");
                ws.Cells["C1"].Insert(eShiftTypeInsert.Right);
            }
        }
        [TestMethod]
        public void ValidateInsertIntoTablePartialDownShouldNotThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("TableDelete");
                ws.Tables.Add(ws.Cells["B2:D3"], "table1");
                ws.Cells["A2"].Insert(eShiftTypeInsert.Down);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ValidateInsertIntoPivotTablePartialRightThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("PivotTableInsert");
                ws.Cells["E5"].Value = "E5";
                ws.Cells["F5"].Value = "F5";
                ws.PivotTables.Add(ws.Cells["B2:D3"], ws.Cells["E5:F6"], "table1");
                ws.Cells["A2"].Insert(eShiftTypeInsert.Right);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ValidateInsertIntoPivotTablePartialDownThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("PivotTableInsert");
                ws.Cells["E5"].Value = "E5";
                ws.Cells["F5"].Value = "F5";
                ws.PivotTables.Add(ws.Cells["B2:D3"], ws.Cells["E5:F6"], "table1");
                ws.Cells["C1"].Insert(eShiftTypeInsert.Down);
            }
        }
        [TestMethod]
        public void ValidateInsertIntoPivotTablePartialRightShouldNotThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("PivotTableInsert");
                ws.Cells["E5"].Value = "E5";
                ws.Cells["F5"].Value = "F5";
                ws.PivotTables.Add(ws.Cells["B2:D3"], ws.Cells["E5:F6"], "table1");
                ws.Cells["C1"].Insert(eShiftTypeInsert.Right);
            }
        }
        [TestMethod]
        public void ValidateInsertIntoPivotTablePartialDownShouldNotThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("PivotTableInsert");
                ws.Cells["E5"].Value = "E5";
                ws.Cells["F5"].Value = "F5";
                ws.PivotTables.Add(ws.Cells["B2:D3"], ws.Cells["E5:F6"], "table1");
                ws.Cells["A2"].Insert(eShiftTypeInsert.Down);
            }
        }
        [TestMethod]
        public void ValidateInsertTableShouldShiftDown()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("TableInsertShiftDown");
                var tbl=ws.Tables.Add(ws.Cells["B2:D3"], "table1");
                ws.Cells["B2:D2"].Insert(eShiftTypeInsert.Down);
                Assert.AreEqual("B3:D4", tbl.Address.Address);

                ws.Cells["A3:D3"].Insert(eShiftTypeInsert.Down);
                Assert.AreEqual("B4:D5", tbl.Address.Address);

                ws.Cells["B3:E3"].Insert(eShiftTypeInsert.Down);
                Assert.AreEqual("B5:D6", tbl.Address.Address);

                //Insert Into
                ws.Cells["B6:E6"].Insert(eShiftTypeInsert.Down);
                Assert.AreEqual("B5:D7", tbl.Address.Address);

                ws.Cells["A6:E6"].Insert(eShiftTypeInsert.Down);
                Assert.AreEqual("B5:D8", tbl.Address.Address);

                ws.Cells["B8:F8"].Insert(eShiftTypeInsert.Down);
                Assert.AreEqual("B5:D9", tbl.Address.Address);
            }
        }
        [TestMethod]
        public void ValidateInsertTableShouldShiftRight()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("TableInsertShiftRight");
                var tbl = ws.Tables.Add(ws.Cells["B2:C4"], "table1");
                ws.Cells["B2:B4"].Insert(eShiftTypeInsert.Right);
                Assert.AreEqual("C2:D4", tbl.Address.Address);

                ws.Cells["B1:B4"].Insert(eShiftTypeInsert.Right);
                Assert.AreEqual("D2:E4", tbl.Address.Address);

                ws.Cells["B2:B6"].Insert(eShiftTypeInsert.Right);
                Assert.AreEqual("E2:F4", tbl.Address.Address);
            }
        }
        [TestMethod]
        public void ValidateInsertPivotTableShouldShiftDown()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("PivotTableInsertShiftDown");
                ws.Cells["E5"].Value = "E5";
                ws.Cells["F5"].Value = "F5";                
                var pt=ws.PivotTables.Add(ws.Cells["B2:D3"], ws.Cells["E5:F6"], "pivottable1");
                ws.Cells["B2:D2"].Insert(eShiftTypeInsert.Down);
                Assert.AreEqual("B3:D4", pt.Address.Address);

                ws.Cells["A2:E2"].Insert(eShiftTypeInsert.Down);
                Assert.AreEqual("B4:D5", pt.Address.Address);

                ws.Cells["B6:D7"].Insert(eShiftTypeInsert.Down);
                Assert.AreEqual("B4:D5", pt.Address.Address);
            }
        }
        [TestMethod]
        public void ValidateInsertPivotTableShouldShiftRight()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("PivotTableInsertShiftRight");
                ws.Cells["E5"].Value = "E5";
                ws.Cells["F5"].Value = "F5";
                var pt = ws.PivotTables.Add(ws.Cells["B2:D3"], ws.Cells["E5:F6"], "pivottable1");
                ws.Cells["B2:B3"].Insert(eShiftTypeInsert.Right);
                Assert.AreEqual("C2:E3", pt.Address.Address);
                ws.Cells["B1:B4"].Insert(eShiftTypeInsert.Right);
                Assert.AreEqual("D2:F3", pt.Address.Address);
                ws.Cells["G2:G3"].Insert(eShiftTypeInsert.Right);
                Assert.AreEqual("D2:F3", pt.Address.Address);
            }
        }

        [TestMethod]
        public void ValidateStyleShiftDown()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("StyleShiftDown");
                ws.Cells["A1"].Style.Fill.SetBackground(OfficeOpenXml.Style.ExcelIndexedColor.Indexed2);
                ws.Cells["B1"].Style.Fill.SetBackground(OfficeOpenXml.Style.ExcelIndexedColor.Indexed3);
                ws.Cells["C1"].Style.Fill.SetBackground(OfficeOpenXml.Style.ExcelIndexedColor.Indexed4);

                ws.Cells["A2"].Style.Fill.SetBackground(OfficeOpenXml.Style.ExcelIndexedColor.Indexed5);
                ws.Cells["A3"].Style.Fill.SetBackground(OfficeOpenXml.Style.ExcelIndexedColor.Indexed6);

                ws.Cells["A1:C1"].Insert(eShiftTypeInsert.Down);
                Assert.AreEqual(0, ws.Cells["A1"].StyleID);
                Assert.AreEqual(0, ws.Cells["B1"].StyleID);
                Assert.AreEqual(0, ws.Cells["C1"].StyleID);
                Assert.AreEqual(2, ws.Cells["A2"].StyleID);
                Assert.AreEqual(3, ws.Cells["B2"].StyleID);
                Assert.AreEqual(4, ws.Cells["C2"].StyleID);
                ws.Cells["A3:C4"].Insert(eShiftTypeInsert.Down);
                Assert.AreEqual(2, ws.Cells["A3"].StyleID);
                Assert.AreEqual(3, ws.Cells["B3"].StyleID);
                Assert.AreEqual(4, ws.Cells["C3"].StyleID);
                Assert.AreEqual(2, ws.Cells["A4"].StyleID);
                Assert.AreEqual(3, ws.Cells["B4"].StyleID);
                Assert.AreEqual(4, ws.Cells["C4"].StyleID);
            }
        }
        [TestMethod]
        public void ValidateStyleShiftRight()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("StyleShiftRight");
                ws.Cells["A1"].Style.Fill.SetBackground(OfficeOpenXml.Style.ExcelIndexedColor.Indexed2);
                ws.Cells["B1"].Style.Fill.SetBackground(OfficeOpenXml.Style.ExcelIndexedColor.Indexed3);
                ws.Cells["C1"].Style.Fill.SetBackground(OfficeOpenXml.Style.ExcelIndexedColor.Indexed4);

                ws.Cells["A2"].Style.Fill.SetBackground(OfficeOpenXml.Style.ExcelIndexedColor.Indexed5);
                ws.Cells["A3"].Style.Fill.SetBackground(OfficeOpenXml.Style.ExcelIndexedColor.Indexed6);

                ws.Cells["A1:A3"].Insert(eShiftTypeInsert.Right);
                Assert.AreEqual(0, ws.Cells["A1"].StyleID);
                Assert.AreEqual(0, ws.Cells["A2"].StyleID);
                Assert.AreEqual(0, ws.Cells["A3"].StyleID);
                Assert.AreEqual(2, ws.Cells["B1"].StyleID);
                Assert.AreEqual(5, ws.Cells["B2"].StyleID);
                Assert.AreEqual(6, ws.Cells["B3"].StyleID);
                ws.Cells["C1:D3"].Insert(eShiftTypeInsert.Right);
                Assert.AreEqual(2, ws.Cells["C1"].StyleID);
                Assert.AreEqual(5, ws.Cells["C2"].StyleID);
                Assert.AreEqual(6, ws.Cells["C3"].StyleID);
                Assert.AreEqual(2, ws.Cells["D1"].StyleID);
                Assert.AreEqual(5, ws.Cells["D2"].StyleID);
                Assert.AreEqual(6, ws.Cells["D3"].StyleID);
            }
        }
        #region Data validation
        [TestMethod]
        public void ValidateDatavalidationFullShiftDown()
        {
            var ws = _pck.Workbook.Worksheets.Add("DataValShiftDownFull");
            var any = ws.DataValidations.AddAnyValidation("B2:E5");

            ws.Cells["A2:E2"].Insert(eShiftTypeInsert.Down);

            Assert.AreEqual("B3:E6", any.Address.Address);
        }
        [TestMethod]
        public void ValidateDatavalidationPartialShiftDown_Left()
        {
            var ws = _pck.Workbook.Worksheets.Add("DataValPartialDownFullL");
            var any = ws.DataValidations.AddAnyValidation("B2:E5");

            ws.Cells["A2:C2"].Insert(eShiftTypeInsert.Down);

            Assert.AreEqual("B3:C6,D2:E5", any.Address.Address);
        }
        [TestMethod]
        public void ValidateDatavalidationPartialShiftDown_Inside()
        {
            var ws = _pck.Workbook.Worksheets.Add("DataValPartialDownFullI");
            var any = ws.DataValidations.AddAnyValidation("B2:E5");

            ws.Cells["C2:D2"].Insert(eShiftTypeInsert.Down);

            Assert.AreEqual("B2:B5,C3:D6,E2:E5", any.Address.Address);
        }


        [TestMethod]
        public void ValidateDatavalidationPartialShiftDown_Right()
        {
            var ws = _pck.Workbook.Worksheets.Add("DataValPartialRightFullR");
            var any = ws.DataValidations.AddAnyValidation("B2:E5");

            ws.Cells["C2:E3"].Insert(eShiftTypeInsert.Down);

            Assert.AreEqual("B2:B5,C4:E7", any.Address.Address);
        }
        [TestMethod]
        public void ValidateDatavalidationPartialShiftRight_Top()
        {
            var ws = _pck.Workbook.Worksheets.Add("DataValPartialRightFullTop");
            var any = ws.DataValidations.AddAnyValidation("B2:E5");

            ws.Cells["A2:A4"].Insert(eShiftTypeInsert.Right);

            Assert.AreEqual("C2:F4,B5:E5", any.Address.Address);
        }
        [TestMethod]
        public void ValidateDatavalidationPartialShiftRight_Inside()
        {
            var ws = _pck.Workbook.Worksheets.Add("DataValPartialRightFullIns");
            var any = ws.DataValidations.AddAnyValidation("B2:E5");

            ws.Cells["A3:A4"].Insert(eShiftTypeInsert.Right);

            Assert.AreEqual("B2:E2,C3:F4,B5:E5", any.Address.Address);
        }
        [TestMethod]
        public void ValidateDatavalidationShiftRight_Inside()
        {
            var ws = _pck.Workbook.Worksheets.Add("dvright");
            var any = ws.DataValidations.AddAnyValidation("B2");

            ws.Cells["B2:C5"].Insert(eShiftTypeInsert.Right);

            Assert.AreEqual("D2", ws.DataValidations[0].Address.Address);
        }

        [TestMethod]
        public void ValidateDatavalPartialShiftRight_Bottom()
        {
            var ws = _pck.Workbook.Worksheets.Add("DataValPartialDownFullBottom");
            var any = ws.DataValidations.AddAnyValidation("B2:E5");

            ws.Cells["A3:A6"].Insert(eShiftTypeInsert.Right);

            Assert.AreEqual("B2:E2,C3:F5", any.Address.Address);
        }

        [TestMethod]
        public void ValidateDatavalidationFullShiftRight()
        {
            var ws = _pck.Workbook.Worksheets.Add("DataValidationShiftRightFull");
            var any = ws.DataValidations.AddAnyValidation("B2:E5");

            ws.Cells["A2:A5"].Insert(eShiftTypeInsert.Right);

            Assert.AreEqual("C2:F5", any.Address.Address);
        }
        [TestMethod]
        public void CheckDataValidationFormulaAfterInsertingRow()
        {
            using (var p = new ExcelPackage())
            {
                // Create a worksheet with conditional formatting 
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var dv = ws.DataValidations.AddCustomValidation("B5:G5");
                dv.Formula.ExcelFormula = "=(B$4=0)";

                // Insert a row before the column being referenced by the CF formula
                ws.InsertRow(2, 1);

                // Check the conditional formatting formula has been updated
                dv = ws.DataValidations[0].As.CustomValidation;
                Assert.AreEqual("=(B$5=0)", dv.Formula.ExcelFormula);
            }
        }
        [TestMethod]
        public void CheckDataValidationFormulaAfterInsertingColumn()
        {
            using (var p = new ExcelPackage())
            {
                // Create a worksheet with conditional formatting 
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var dv = ws.DataValidations.AddCustomValidation("E2:E7");
                dv.Formula.ExcelFormula = "=($D2=0)";

                // Insert a column before the column being referenced by the CF formula
                ws.InsertColumn(2, 1);

                // Check the conditional formatting formula has been updated
                dv = ws.DataValidations[0].As.CustomValidation;
                Assert.AreEqual("=($E2=0)", dv.Formula.ExcelFormula);
            }
        }
        #endregion
        #region Conditional formatting
        [TestMethod]
        public void ValidateConditionalFormattingFullShiftDown()
        {
            var ws = _pck.Workbook.Worksheets.Add("CondFormShiftDownFull");
            var cf = ws.ConditionalFormatting.AddAboveAverage(new ExcelAddress("B2:E5"));
            cf.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent1);
            ws.Cells["A2:E2"].Insert(eShiftTypeInsert.Down);

            Assert.AreEqual("B3:E6", cf.Address.Address);
        }
        [TestMethod]
        public void ValidateConditionalFormattingPartialShiftDown_Left()
        {
            var ws = _pck.Workbook.Worksheets.Add("CondFormPartialDownFullL");
            var cf = ws.ConditionalFormatting.AddAboveAverage(new ExcelAddress("B2:E5"));
            cf.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent1);

            ws.Cells["A2:C2"].Insert(eShiftTypeInsert.Down);

            Assert.AreEqual("B3:C6,D2:E5", cf.Address.Address);
        }
        [TestMethod]
        public void ValidateConditionalFormattingShiftDown_Inside()
        {
            var ws = _pck.Workbook.Worksheets.Add("CondFormPartialDownFullI");
            var cf = ws.ConditionalFormatting.AddAboveAverage(new ExcelAddress("B2:E5"));
            cf.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent1);

            ws.Cells["C2:D2"].Insert(eShiftTypeInsert.Down);

            Assert.AreEqual("B2:B5,C3:D6,E2:E5", cf.Address.Address);
        }


        [TestMethod]
        public void ValidateConditionalFormattingShiftDown_Right()
        {
            var ws = _pck.Workbook.Worksheets.Add("CondFormPartialRightFullR");
            var cf = ws.ConditionalFormatting.AddAboveAverage(new ExcelAddress("B2:E5"));
            cf.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent1);

            ws.Cells["C2:E3"].Insert(eShiftTypeInsert.Down);

            Assert.AreEqual("B2:B5,C4:E7", cf.Address.Address);
        }
        [TestMethod]
        public void ValidateConditionalFormattingPartialShiftRight_Top()
        {
            var ws = _pck.Workbook.Worksheets.Add("CondFormPartialRightFullTop");
            var cf = ws.ConditionalFormatting.AddAboveAverage(new ExcelAddress("B2:E5"));
            cf.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent1);

            ws.Cells["A2:A4"].Insert(eShiftTypeInsert.Right);

            Assert.AreEqual("C2:F4,B5:E5", cf.Address.Address);
        }
        [TestMethod]
        public void ValidateConditionalFormattingPartialShiftRight_Inside()
        {
            var ws = _pck.Workbook.Worksheets.Add("CondFormPartialRightFullIns");
            var cf = ws.ConditionalFormatting.AddAboveAverage(new ExcelAddress("B2:E5"));
            cf.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent1);

            ws.Cells["A3:A4"].Insert(eShiftTypeInsert.Right);

            Assert.AreEqual("B2:E2,C3:F4,B5:E5", cf.Address.Address);
        }

        [TestMethod]
        public void ValidateConditionalFormattingShiftRight_Bottom()
        {
            var ws = _pck.Workbook.Worksheets.Add("CondFormPartialDownFullBottom");
            var cf = ws.ConditionalFormatting.AddAboveAverage(new ExcelAddress("B2:E5"));
            cf.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent1);

            ws.Cells["A3:A6"].Insert(eShiftTypeInsert.Right);

            Assert.AreEqual("B2:E2,C3:F5", cf.Address.Address);
        }

        [TestMethod]
        public void ValidateConditionalFormattingFullShiftRight()
        {
            var ws = _pck.Workbook.Worksheets.Add("CondFormShiftRightFull");
            var cf = ws.ConditionalFormatting.AddAboveAverage(new ExcelAddress("B2:E5"));
            cf.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent1);

            ws.Cells["A2:A5"].Insert(eShiftTypeInsert.Right);

            Assert.AreEqual("C2:F5", cf.Address.Address);
        }
        #endregion
        [TestMethod]
        public void ValidateFilterShiftDown()
        {
            var ws = _pck.Workbook.Worksheets.Add("AutoFilterShiftDown");
            LoadTestdata(ws);
            ws.AutoFilterAddress = new ExcelAddressBase("A1:D100");
            ws.Cells["A1:D1"].Insert(eShiftTypeInsert.Down);
            Assert.AreEqual("A2:D101", ws.AutoFilterAddress.Address);
            ws.Cells["A50:D50"].Insert(eShiftTypeInsert.Down);
            Assert.AreEqual("A2:D102", ws.AutoFilterAddress.Address);
        }
        [TestMethod]
        public void ValidateFilterShiftRight()
        {
            var ws = _pck.Workbook.Worksheets.Add("AutoFilterShiftRight");
            LoadTestdata(ws);
            ws.AutoFilterAddress = new ExcelAddressBase("A1:D100");
            ws.Cells["A1:A100"].Insert(eShiftTypeInsert.Right);
            Assert.AreEqual("B1:E100", ws.AutoFilterAddress.Address);
            ws.Cells["C1:C100"].Insert(eShiftTypeInsert.Right);
            Assert.AreEqual("B1:F100", ws.AutoFilterAddress.Address);
        }
        [TestMethod]
        public void ValidateFilterInsertRow()
        {
            var ws = _pck.Workbook.Worksheets.Add("AutoFilterInsertRow");
            LoadTestdata(ws);
            ws.AutoFilterAddress = new ExcelAddressBase("A1:D100");
            ws.InsertRow(1,1);
            Assert.AreEqual("A2:D101", ws.AutoFilterAddress.Address);
            ws.InsertRow(5, 2);
            Assert.AreEqual("A2:D103", ws.AutoFilterAddress.Address);
        }
        [TestMethod]
        public void ValidateFilterInsertColumn()
        {
            var ws = _pck.Workbook.Worksheets.Add("AutoFilterInsertCol");
            LoadTestdata(ws);
            ws.AutoFilterAddress = new ExcelAddressBase("A1:D100");
            ws.InsertColumn(1,1);
            Assert.AreEqual("B1:E100", ws.AutoFilterAddress.Address);
            ws.InsertColumn(3,2);
            Assert.AreEqual("B1:G100", ws.AutoFilterAddress.Address);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ValidateFilterShiftDownPartial()
        {
            var ws = _pck.Workbook.Worksheets.Add("AutoFilterShiftDownPart");
            LoadTestdata(ws);
            ws.AutoFilterAddress = new ExcelAddressBase("A1:D100");
            ws.Cells["A1:C1"].Insert(eShiftTypeInsert.Down);
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ValidateFilterShiftRightPartial()
        {
            var ws = _pck.Workbook.Worksheets.Add("AutoFilterShiftRightPart");
            LoadTestdata(ws);
            ws.AutoFilterAddress = new ExcelAddressBase("A1:D100");
            ws.Cells["A1:A99"].Insert(eShiftTypeInsert.Right);
        }
        [TestMethod]
        public void ValidateSparkLineShiftRight()
        {
            var ws = _pck.Workbook.Worksheets.Add("SparkLineShiftRight");
            LoadTestdata(ws,10);
            ws.SparklineGroups.Add(OfficeOpenXml.Sparkline.eSparklineType.Column, ws.Cells["E2:E10"], ws.Cells["A2:D10"]);
            ws.Cells["E5"].Insert(eShiftTypeInsert.Right);
            Assert.AreEqual("F5", ws.SparklineGroups[0].Sparklines[3].Cell.Address);
            ws.Cells["A1:A10"].Insert(eShiftTypeInsert.Right);
            Assert.AreEqual("B2:E10", ws.SparklineGroups[0].DataRange.Address);
        }
        [TestMethod]
        public void ValidateSparkLineShiftDown()
        {
            var ws = _pck.Workbook.Worksheets.Add("SparkLineShiftDown");
            LoadTestdata(ws, 10);
            ws.SparklineGroups.Add(OfficeOpenXml.Sparkline.eSparklineType.Column, ws.Cells["E2:E10"], ws.Cells["A2:D10"]);
            ws.Cells["E5"].Insert(eShiftTypeInsert.Down);
            Assert.AreEqual("E6", ws.SparklineGroups[0].Sparklines[3].Cell.Address);
            ws.Cells["A1:E1"].Insert(eShiftTypeInsert.Down);
            Assert.AreEqual("A3:D11", ws.SparklineGroups[0].DataRange.Address);
        }
        [TestMethod]
        public void ValidateSparkLineInsertRow()
        {
            var ws = _pck.Workbook.Worksheets.Add("SparkLineInsertRow");
            LoadTestdata(ws, 10);
            ws.SparklineGroups.Add(OfficeOpenXml.Sparkline.eSparklineType.Column, ws.Cells["E2:E10"], ws.Cells["A2:D10"]);
            ws.InsertRow(5, 1);
            Assert.AreEqual("E6", ws.SparklineGroups[0].Sparklines[3].Cell.Address);
            ws.InsertRow(1, 1);
            Assert.AreEqual("A3:D12", ws.SparklineGroups[0].DataRange.Address);
        }
        [TestMethod]
        public void ValidateSparkLineInsertColumn()
        {
            var ws = _pck.Workbook.Worksheets.Add("SparkLineInsertColumn");
            LoadTestdata(ws, 10);
            ws.SparklineGroups.Add(OfficeOpenXml.Sparkline.eSparklineType.Column, ws.Cells["E2:E10"], ws.Cells["A2:D10"]);
            ws.InsertColumn(2, 1);
            Assert.AreEqual("F5", ws.SparklineGroups[0].Sparklines[3].Cell.Address);
            ws.InsertColumn(1, 1);
            Assert.AreEqual("B2:F10", ws.SparklineGroups[0].DataRange.Address);
        }

        [TestMethod]
        public void InsertIntoTemplate1()
        {
            using (var p = OpenTemplatePackage("InsertDeleteTemplate.xlsx"))
            {
                var ws = p.Workbook.Worksheets["C3R"];
                var ws2 = ws.Workbook.Worksheets.Add("C3R-2", ws);
                ws.Cells["G49:G52"].Insert(eShiftTypeInsert.Down);
                ws2.Cells["G49:G52"].Insert(eShiftTypeInsert.Right);

                SaveWorkbook("InsertTest1.xlsx", p);
            }
        }
        [TestMethod]
        public void InsertIntoTemplate2()
        {
            using (var p = OpenTemplatePackage("InsertDeleteTemplate.xlsx"))
            {
                var ws = p.Workbook.Worksheets["C3R"];
                var ws2 = ws.Workbook.Worksheets.Add("C3R-2", ws);
                ws.Cells["L49:L52"].Insert(eShiftTypeInsert.Down);
                ws2.Cells["L49:L52"].Insert(eShiftTypeInsert.Right);

                SaveWorkbook("InsertTest2.xlsx", p);
            }
        }
        [TestMethod]
        public void ValidateConditionalFormattingInsertColumnMultiRange()
        {
            var ws = _pck.Workbook.Worksheets.Add("CondFormPartialUpMR");
            var cf = ws.ConditionalFormatting.AddAboveAverage(new ExcelAddress("B2:E5,D3:E5"));
            cf.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent1);

            ws.InsertColumn(4,1);

            Assert.AreEqual("B2:F5,E3:F5", cf.Address.Address);
        }
        [TestMethod]
        public void CheckConditionalFormattingFormulaAfterInsertingRow()
        {
            using (var p = new ExcelPackage())
            {
                // Create a worksheet with conditional formatting 
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var cf = ws.ConditionalFormatting.AddExpression(ws.Cells["B5:G5"]);
                cf.Formula = "=(B$4=0)";

                // Insert a row before the column being referenced by the CF formula
                ws.InsertRow(2, 1);

                // Check the conditional formatting formula has been updated
                cf = ws.ConditionalFormatting[0].As.Expression;
                Assert.AreEqual("=(B$5=0)", cf.Formula);
            }
        }
        [TestMethod]
        public void CheckConditionalFormattingFormulaAfterInsertingColumn()
        {
            using (var p = new ExcelPackage())
            {
                // Create a worksheet with conditional formatting 
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var cf = ws.ConditionalFormatting.AddExpression(ws.Cells["E2:E7"]);
                cf.Formula = "=($D2=0)";

                // Insert a column before the column being referenced by the CF formula
                ws.InsertColumn(2, 1);

                // Check the conditional formatting formula has been updated
                cf = ws.ConditionalFormatting[0].As.Expression;
                Assert.AreEqual("=($E2=0)", cf.Formula);
            }
        }

        [TestMethod]
        public void ValidateCommentsShouldShiftRightOnInsertIntoRange()
        {
            var ws = _pck.Workbook.Worksheets.Add("InsertRightComment");
            var commentAddress = "B2";
            ws.Comments.Add(ws.Cells[commentAddress], "This is a comment.", "author");
            ws.Cells[commentAddress].Value = "This cell contains a comment.";

            ws.Cells["B1:B3"].Insert(eShiftTypeInsert.Right);
            commentAddress = "C2";

            Assert.AreEqual(1, ws.Comments.Count);
            Assert.AreEqual("This is a comment.", ws.Comments[0].Text);
            Assert.AreEqual("This cell contains a comment.", ws.Cells[commentAddress].GetValue<string>());
            Assert.AreEqual(commentAddress, ws.Comments[0].Address);
        }
        [TestMethod]
        public void ValidateCommentsShouldNotShiftRightOnInsertIntoRange()
        {
            var ws = _pck.Workbook.Worksheets.Add("InsertRightComment2");
            var commentAddress = "B4";
            ws.Comments.Add(ws.Cells[commentAddress], "This is a comment.", "author");
            ws.Cells[commentAddress].Value = "This cell contains a comment.";

            ws.Cells["B1:B3"].Insert(eShiftTypeInsert.Right);

            Assert.AreEqual(1, ws.Comments.Count);
            Assert.AreEqual("This is a comment.", ws.Comments[0].Text);
            Assert.AreEqual("This cell contains a comment.", ws.Cells[commentAddress].GetValue<string>());
            Assert.AreEqual(commentAddress, ws.Comments[0].Address);
        }
        [TestMethod]
        public void ValidateThreadedCommentsShouldShiftRightOnInsertIntoRange()
        {
            var ws = _pck.Workbook.Worksheets.Add("InsertRightTC");
            var commentAddress = "B2";
            ws.ThreadedComments.Add(commentAddress);
            ws.Cells[commentAddress].Value = "This cell contains a threaded comment.";

            ws.Cells["B1:B3"].Insert(eShiftTypeInsert.Right);
            commentAddress = "C2";

            Assert.AreEqual(1, ws.ThreadedComments.Count);
            Assert.AreEqual("This cell contains a threaded comment.", ws.Cells[commentAddress].GetValue<string>());
            Assert.AreEqual(commentAddress, ws.ThreadedComments[0].CellAddress.Address);
        }
        [TestMethod]
        public void ValidateThreadedCommentsShouldNotShiftRightOnInsertIntoRange()
        {
            var ws = _pck.Workbook.Worksheets.Add("InsertRightTC2");
            var commentAddress = "B4";
            ws.ThreadedComments.Add(commentAddress);
            ws.Cells[commentAddress].Value = "This cell contains a threaded comment.";

            ws.Cells["B1:B3"].Insert(eShiftTypeInsert.Right);

            Assert.AreEqual(1, ws.ThreadedComments.Count);
            Assert.AreEqual("This cell contains a threaded comment.", ws.Cells[commentAddress].GetValue<string>());
            Assert.AreEqual(commentAddress, ws.ThreadedComments[0].CellAddress.Address);
        }
        [TestMethod]
        public void ValidateTableCalculatedColumnFormulasAfterInsertRowAndInsertColumn()
        {
            //Test created from issue #484 - https://github.com/EPPlusSoftware/EPPlus/issues/484
            var ws = _pck.Workbook.Worksheets.Add("InsertCalculateColumnFormula");

            // Create some tables with calculated column formulas
            var tbl1 = ws.Tables.Add(ws.Cells["A11:C15"], "Table3");
            tbl1.Columns[2].CalculatedColumnFormula = "A12+B12";

            var tbl2 = ws.Tables.Add(ws.Cells["E11:G15"], "Table4");
            tbl2.Columns[2].CalculatedColumnFormula = "A12+F12";

            // Check the formulas have been set correctly
            Assert.AreEqual("A12+B12", ws.Cells["C12"].Formula);
            Assert.AreEqual("A12+F12", ws.Cells["G12"].Formula);
            Assert.AreEqual("A12+B12", tbl1.Columns[2].CalculatedColumnFormula);
            Assert.AreEqual("A12+F12", tbl2.Columns["Column3"].CalculatedColumnFormula);

            // Insert two rows above the tables
            ws.InsertRow(5, 2);
            // Insert one column from column D
            ws.InsertColumn(4, 1);

            // Check the formulas were updated
            Assert.AreEqual("A14+B14", ws.Cells["C14"].Formula);
            Assert.AreEqual("A14+G14", ws.Cells["H14"].Formula);
            Assert.AreEqual("A15+G15", ws.Cells["H15"].Formula);
            Assert.AreEqual("A14+B14", tbl1.Columns[2].CalculatedColumnFormula);
            Assert.AreEqual("A14+G14", tbl2.Columns[2].CalculatedColumnFormula);
        }
        [TestMethod]
        public void ValidateTableCalculatedColumnFormulasAfterInsertRange()
        {
            //Test created from issue #484 - https://github.com/EPPlusSoftware/EPPlus/issues/484
            var ws = _pck.Workbook.Worksheets.Add("InsertCalcColumnFormulaRange");

            // Create some tables with calculated column formulas
            var tbl1 = ws.Tables.Add(ws.Cells["A11:C15"], "Table1");
            tbl1.Columns[2].CalculatedColumnFormula = "A12+B12";

            var tbl2 = ws.Tables.Add(ws.Cells["E11:G15"], "Table2");
            tbl2.Columns[2].CalculatedColumnFormula = "A12+F12";

            // Check the formulas have been set correctly
            Assert.AreEqual("A12+B12", ws.Cells["C12"].Formula);
            Assert.AreEqual("A12+F12", ws.Cells["G12"].Formula);
            Assert.AreEqual("A12+B12", tbl1.Columns[2].CalculatedColumnFormula);
            Assert.AreEqual("A12+F12", tbl2.Columns["Column3"].CalculatedColumnFormula);

            // Insert two rows above the tables
            ws.Cells["A2:D2"].Insert(eShiftTypeInsert.Down);
            // Insert one column from column D
            //ws.InsertColumn(4, 1);
            ws.Cells["A1:A20"].Insert(eShiftTypeInsert.Right);

            // Check the formulas were updated
            Assert.AreEqual("B14+C14", ws.Cells["D14"].Formula);
            Assert.AreEqual("B13+G12", ws.Cells["H12"].Formula);
            Assert.AreEqual("B14+G13", ws.Cells["H13"].Formula);
            Assert.AreEqual("B13+C13", tbl1.Columns[2].CalculatedColumnFormula);
            Assert.AreEqual("B13+G12", tbl2.Columns[2].CalculatedColumnFormula);
        }

    }
}