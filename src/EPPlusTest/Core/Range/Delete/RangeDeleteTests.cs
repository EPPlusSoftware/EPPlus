using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EPPlusTest.Core.Range.Delete
{
    [TestClass]
    public class RangeDeleteTests : TestBase
    {
        public static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("WorksheetRangeDelete.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void ValidateFormulasAfterDeleteRow()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DeleteRow_Sheet1");
            var ws2 = _pck.Workbook.Worksheets.Add("DeleteRow_Sheet2");
            ws.Cells["A1"].Formula = "Sum(C5:C10)";
            ws.Cells["B1:B2"].Formula = "Sum(C5:C10)";
            ws2.Cells["A1"].Formula = "Sum(DeleteRow_Sheet1!C5:C10)";
            ws2.Cells["B1:B2"].Formula = "Sum(DeleteRow_Sheet1!C5:C10)";

            //Act
            ws.DeleteRow(3, 1);
            var wsError = _pck.Workbook.Worksheets["DeleteRow_Sheet1"];
            if (wsError != null)
            {
                Assert.AreEqual(1, wsError._sharedFormulas.Count);
            }


            //Assert
            Assert.AreEqual("Sum(C4:C9)", ws.Cells["A1"].Formula);
            Assert.AreEqual("Sum(C4:C9)", ws.Cells["B1"].Formula);
            Assert.AreEqual("Sum(C5:C10)", ws.Cells["B2"].Formula);

            Assert.AreEqual("Sum(DeleteRow_Sheet1!C4:C9)", ws2.Cells["A1"].Formula);
            Assert.AreEqual("Sum(DeleteRow_Sheet1!C4:C9)", ws2.Cells["B1"].Formula);
            Assert.AreEqual("Sum(DeleteRow_Sheet1!C5:C10)", ws2.Cells["B2"].Formula);
        }
        [TestMethod]
        public void ValidateFormulasAfterDelete2Rows()
        {
            //Setup
            var ws1 = _pck.Workbook.Worksheets.Add("DeleteRow2_Sheet1");
            var ws2 = _pck.Workbook.Worksheets.Add("DeleteRow2_Sheet2");
            ws1.Cells["B3:B6"].Formula = "A1+C3";
            ws2.Cells["B3:B6"].Formula = "DeleteRow2_Sheet1!A1+DeleteRow2_Sheet1!C2";

            //Act
            ws1.DeleteRow(2, 2);
            var wsError = _pck.Workbook.Worksheets["DeleteRow_Sheet1"];
            if (wsError != null)
            {
                Assert.AreEqual(1, wsError._sharedFormulas.Count);
            }

            //Assert
            Assert.AreEqual("", ws1.Cells["B1"].Formula);
            Assert.AreEqual("#REF!+C2", ws1.Cells["B2"].Formula);
            Assert.AreEqual("#REF!+C3", ws1.Cells["B3"].Formula);
            Assert.AreEqual("A2+C4", ws1.Cells["B4"].Formula);
            Assert.AreEqual("", ws1.Cells["B6"].Formula);

            Assert.AreEqual("DeleteRow2_Sheet1!A1+DeleteRow2_Sheet1!#REF!", ws2.Cells["B3"].Formula);
            Assert.AreEqual("DeleteRow2_Sheet1!#REF!+DeleteRow2_Sheet1!#REF!", ws2.Cells["B4"].Formula);
            Assert.AreEqual("DeleteRow2_Sheet1!#REF!+DeleteRow2_Sheet1!C2", ws2.Cells["B5"].Formula);
            Assert.AreEqual("DeleteRow2_Sheet1!A2+DeleteRow2_Sheet1!C3", ws2.Cells["B6"].Formula);
        }
        [TestMethod]
        public void ValidateFormulasAfterDeleteColumn()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DeleteCol_Sheet1");
            var ws2 = _pck.Workbook.Worksheets.Add("DeleteCol_Sheet2");
            ws.Cells["A1"].Formula = "Sum(E3:I3)";
            ws.Cells["A2:B2"].Formula = "Sum(E3:I3)";
            ws2.Cells["A1"].Formula = "Sum(DeleteCol_Sheet1!E3:I3)";
            ws2.Cells["A2:B2"].Formula = "Sum(DeleteCol_Sheet1!E3:I3)";

            //Act
            ws.DeleteColumn(3, 1);
            var wsError = _pck.Workbook.Worksheets["DeleteRow_Sheet1"];
            if (wsError != null)
            {
                Assert.AreEqual(1, wsError._sharedFormulas.Count);
            }

            //Assert
            Assert.AreEqual("Sum(D3:H3)", ws.Cells["A1"].Formula);
            Assert.AreEqual("Sum(D3:H3)", ws.Cells["A2"].Formula);
            Assert.AreEqual("Sum(E3:I3)", ws.Cells["B2"].Formula);

            Assert.AreEqual("Sum(DeleteCol_Sheet1!D3:H3)", ws2.Cells["A1"].Formula);
            Assert.AreEqual("Sum(DeleteCol_Sheet1!D3:H3)", ws2.Cells["A2"].Formula);
            Assert.AreEqual("Sum(DeleteCol_Sheet1!E3:I3)", ws2.Cells["B2"].Formula);
        }
        [TestMethod]
        public void ValidateFormulasAfterDelete2Columns()
        {
            //Setup
            var ws1 = _pck.Workbook.Worksheets.Add("DeleteCol2_Sheet1");
            var ws2 = _pck.Workbook.Worksheets.Add("DeleteCol2_Sheet2");
            ws1.Cells["C2:F2"].Formula = "A1+C3";
            ws2.Cells["C2:F2"].Formula = "DeleteCol2_Sheet1!A1+DeleteCol2_Sheet1!C3";

            //Act
            ws1.DeleteColumn(2, 2);
            var wsError = _pck.Workbook.Worksheets["DeleteRow_Sheet1"];
            if (wsError != null)
            {
                Assert.AreEqual(1, wsError._sharedFormulas.Count);
            }

            //Assert
            Assert.AreEqual("", ws1.Cells["A2"].Formula);
            Assert.AreEqual("#REF!+B3", ws1.Cells["B2"].Formula);
            Assert.AreEqual("#REF!+C3", ws1.Cells["C2"].Formula);
            Assert.AreEqual("B1+D3", ws1.Cells["D2"].Formula);
            Assert.AreEqual("", ws1.Cells["B6"].Formula);

            Assert.AreEqual("DeleteCol2_Sheet1!A1+DeleteCol2_Sheet1!#REF!", ws2.Cells["C2"].Formula);
            Assert.AreEqual("DeleteCol2_Sheet1!#REF!+DeleteCol2_Sheet1!B3", ws2.Cells["D2"].Formula);
            Assert.AreEqual("DeleteCol2_Sheet1!#REF!+DeleteCol2_Sheet1!C3", ws2.Cells["E2"].Formula);
            Assert.AreEqual("DeleteCol2_Sheet1!B1+DeleteCol2_Sheet1!D3", ws2.Cells["F2"].Formula);
        }
        [TestMethod]
        public void SharedFormulaShouldBeDeletedIfEntireRowIsDeleted()
        {
            using (var p = new ExcelPackage())
            {
                //Setup
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A2:B2"].Formula = "C2";
                //Act
                Assert.AreEqual(1, ws._sharedFormulas.Count);
                ws.DeleteRow(2);

                //Assert
                Assert.AreEqual(0, ws._sharedFormulas.Count);
                Assert.AreEqual("", ws.Cells["A2"].Formula);
                Assert.AreEqual("", ws.Cells["B2"].Formula);
            }
        }
        [TestMethod]
        public void SharedFormulaShouldBeDeletedIfEntireColumnIsDeleted()
        {
            using (var p = new ExcelPackage())
            {
                //Setup
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["B1:B2"].Formula = "C2";
                //Act
                Assert.AreEqual(1, ws._sharedFormulas.Count);
                ws.DeleteColumn(2);

                //Assert
                Assert.AreEqual(0, ws._sharedFormulas.Count);
                Assert.AreEqual("", ws.Cells["B1"].Formula);
                Assert.AreEqual("", ws.Cells["B2"].Formula);
            }
        }
        [TestMethod]
        public void SharedFormulaShouldBePartialDeletedRow()
        {
            using (var p = new ExcelPackage())
            {
                //Setup
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A2:B3"].Formula = "C2";
                //Act
                Assert.AreEqual(1, ws._sharedFormulas.Count);
                ws.DeleteRow(2);

                //Assert
                Assert.AreEqual(0, ws._sharedFormulas.Count);
                Assert.AreEqual("C2", ws.Cells["A2"].Formula);
                Assert.AreEqual("D2", ws.Cells["B2"].Formula);
                Assert.AreEqual("", ws.Cells["A3"].Formula);
                Assert.AreEqual("", ws.Cells["B3"].Formula);
            }
        }
        [TestMethod]
        public void SharedFormulaShouldBePartialDeletedColumn()
        {
            using (var p = new ExcelPackage())
            {
                //Setup
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["B1:C2"].Formula = "B3";
                //Act
                Assert.AreEqual(1, ws._sharedFormulas.Count);
                ws.DeleteColumn(2);

                //Assert
                Assert.AreEqual(0, ws._sharedFormulas.Count);
                Assert.AreEqual("B3", ws.Cells["B1"].Formula);
                Assert.AreEqual("B4", ws.Cells["B2"].Formula);
                Assert.AreEqual("", ws.Cells["C1"].Formula);
                Assert.AreEqual("", ws.Cells["C2"].Formula);
            }
        }
        [TestMethod]
        public void SharedFormulaShouldBePartialDeletedRowShareFormulaRetained()
        {
            using (var p = new ExcelPackage())
            {
                //Setup
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A2:B3"].Formula = "E12";
                //Act
                Assert.AreEqual(1, ws._sharedFormulas.Count);
                ws.DeleteRow(2);

                //Assert
                Assert.AreEqual(1, ws._sharedFormulas.Count);
                Assert.AreEqual("E11", ws.Cells["A2"].Formula);
                Assert.AreEqual("F11", ws.Cells["B2"].Formula);
                Assert.AreEqual("", ws.Cells["A3"].Formula);
                Assert.AreEqual("", ws.Cells["B3"].Formula);
            }
        }
        [TestMethod]
        public void SharedFormulaShouldBePartialDeletedColumnShareFormulaRetained()
        {
            using (var p = new ExcelPackage())
            {
                //Setup
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["B1:C2"].Formula = "E12";
                //Act
                Assert.AreEqual(1, ws._sharedFormulas.Count);
                ws.DeleteColumn(2);

                //Assert
                Assert.AreEqual(1, ws._sharedFormulas.Count);
                Assert.AreEqual("D12", ws.Cells["B1"].Formula);
                Assert.AreEqual("D13", ws.Cells["B2"].Formula);
                Assert.AreEqual("", ws.Cells["C1"].Formula);
                Assert.AreEqual("", ws.Cells["C2"].Formula);
            }
        }
        [TestMethod]
        public void FixedAddressesShouldBeDeletedRow()
        {
            using (var p = new ExcelPackage())
            {
                //Setup
                var ws1 = p.Workbook.Worksheets.Add("Sheet1");
                var ws2 = p.Workbook.Worksheets.Add("Sheet2");
                ws1.Cells["A1"].Formula = "SUM($A$5:$A$8)";
                ws2.Cells["A1"].Formula = "SUM(sheet1!$A$5:$A$8)";
                //Act
                ws1.DeleteRow(4);
                Assert.AreEqual("SUM($A$4:$A$7)", ws1.Cells["A1"].Formula);
                Assert.AreEqual("SUM(sheet1!$A$4:$A$7)", ws2.Cells["A1"].Formula);
                ws1.DeleteRow(4);
                Assert.AreEqual("SUM($A$4:$A$6)", ws1.Cells["A1"].Formula);
                Assert.AreEqual("SUM(sheet1!$A$4:$A$6)", ws2.Cells["A1"].Formula);
                ws1.DeleteRow(6);
                Assert.AreEqual("SUM($A$4:$A$5)", ws1.Cells["A1"].Formula);
                Assert.AreEqual("SUM(sheet1!$A$4:$A$5)", ws2.Cells["A1"].Formula);
                ws1.DeleteRow(6);
                Assert.AreEqual("SUM($A$4:$A$5)", ws1.Cells["A1"].Formula);
                Assert.AreEqual("SUM(sheet1!$A$4:$A$5)", ws2.Cells["A1"].Formula);
            }
        }
        [TestMethod]
        public void FixedAddressesShouldBeDeletedColumn()
        {
            using (var p = new ExcelPackage())
            {
                //Setup
                var ws1 = p.Workbook.Worksheets.Add("Sheet1");
                var ws2 = p.Workbook.Worksheets.Add("Sheet2");
                ws1.Cells["A1"].Formula = "SUM($E$1:$H$1)";
                ws2.Cells["A1"].Formula = "SUM(sheet1!$E$1:$H$1)";
                //Act
                ws1.DeleteColumn(4);
                Assert.AreEqual("SUM($D$1:$G$1)", ws1.Cells["A1"].Formula);
                Assert.AreEqual("SUM(sheet1!$D$1:$G$1)", ws2.Cells["A1"].Formula);

                ws1.DeleteColumn(4);
                Assert.AreEqual("SUM($D$1:$F$1)", ws1.Cells["A1"].Formula);
                Assert.AreEqual("SUM(sheet1!$D$1:$F$1)", ws2.Cells["A1"].Formula);

                ws1.DeleteColumn(6);
                Assert.AreEqual("SUM($D$1:$E$1)", ws1.Cells["A1"].Formula);
                Assert.AreEqual("SUM(sheet1!$D$1:$E$1)", ws2.Cells["A1"].Formula);

                ws1.DeleteColumn(6);
                Assert.AreEqual("SUM($D$1:$E$1)", ws1.Cells["A1"].Formula);
                Assert.AreEqual("SUM(sheet1!$D$1:$E$1)", ws2.Cells["A1"].Formula);
            }
        }
        [TestMethod]
        public void ValidateValuesAfterDeleteRowInRangeShiftUp()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DeleteRangeDown");
            SetValues(ws,3);

            //Act
            ws.Cells["B2"].Delete(eShiftTypeDelete.Up);

            //Assert
            Assert.AreEqual("A1", ws.Cells["A1"].Value);
            Assert.AreEqual("A2", ws.Cells["A2"].Value);
            Assert.AreEqual("A3", ws.Cells["A3"].Value);
            Assert.AreEqual("B1", ws.Cells["B1"].Value);
            Assert.AreEqual("B3", ws.Cells["B2"].Value);
            Assert.IsNull(ws.Cells["B3"].Value);
            Assert.AreEqual("C1", ws.Cells["C1"].Value);
            Assert.AreEqual("C2", ws.Cells["C2"].Value);
            Assert.AreEqual("C3", ws.Cells["C3"].Value);
        }
        [TestMethod]
        public void ValidateValuesAfterDeleteRowInRangeShiftLeft()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DeleteRangeLeft");
            SetValues(ws, 3);

            //Act
            ws.Cells["B2"].Delete(eShiftTypeDelete.Left);

            //Assert
            Assert.AreEqual("A1", ws.Cells["A1"].Value);
            Assert.AreEqual("A2", ws.Cells["A2"].Value);
            Assert.AreEqual("A3", ws.Cells["A3"].Value);
            Assert.AreEqual("B1", ws.Cells["B1"].Value);
            Assert.AreEqual("C2", ws.Cells["B2"].Value);
            Assert.AreEqual("C1", ws.Cells["C1"].Value);
            Assert.IsNull(ws.Cells["C2"].Value);
            Assert.AreEqual("C3", ws.Cells["C3"].Value);

            //Act 2
            ws.Cells["A1:B1"].Delete(eShiftTypeDelete.Left);
            
            //Assert 2
            Assert.AreEqual("C1", ws.Cells["A1"].Value);
            Assert.IsNull(ws.Cells["B1"].Value);
            Assert.IsNull(ws.Cells["C1"].Value);
        }

        [TestMethod]
        public void ValidateValuesAfterDeleteInRangeShiftUpTwoRows()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DeleteRangeUpTwoRows");
            SetValues(ws, 4);

            //Act
            ws.Cells["B1:C2"].Delete(eShiftTypeDelete.Up);

            //Assert
            AssertNoChange(ws.Cells["A1:A4,D1:D4"]);
            AssertIsNull(ws.Cells["B3:C4"]);

            Assert.AreEqual("B3", ws.Cells["B1"].Value);
            Assert.AreEqual("B4", ws.Cells["B2"].Value);            
            Assert.AreEqual("C3", ws.Cells["C1"].Value);
            Assert.AreEqual("C4", ws.Cells["C2"].Value);
        }
        [TestMethod]
        public void ValidateValuesAfterDeleteInRangeShiftLeftTwoRows()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DeleteRangeLeftTwoRows");
            SetValues(ws, 4);

            //Act
            ws.Cells["B1:C2"].Delete(eShiftTypeDelete.Left);

            //Assert
            AssertNoChange(ws.Cells["A1:A4,D1:D4"]);
            AssertIsNull(ws.Cells["C1:D2"]);

            Assert.AreEqual("D1", ws.Cells["B1"].Value);
            Assert.AreEqual("D2", ws.Cells["B2"].Value);
        }


        [TestMethod]
        public void ValidateCommentsAfterDeleteShiftUp()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DeleteRangeCommentsUp");
            ws.Cells["A1"].AddComment("Comment A1", "EPPlus");
            ws.Cells["A2"].AddComment("Comment A2", "EPPlus");
            ws.Cells["A3"].AddComment("Comment A3", "EPPlus");

            //Act
            ws.Cells["A2"].Delete(eShiftTypeDelete.Up);

            //Assert
            Assert.AreEqual("Comment A1", ws.Cells["A1"].Comment.Text);
            Assert.AreEqual("Comment A3", ws.Cells["A2"].Comment.Text);
            Assert.IsNull(ws.Cells["A3"].Comment);
        }
        [TestMethod]
        public void ValidateCommentsAfterDeleteShiftLeft()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DeleteRangeCommentsLeft");
            ws.Cells["A1"].AddComment("Comment A1", "EPPlus");
            ws.Cells["B1"].AddComment("Comment B1", "EPPlus");
            ws.Cells["C1"].AddComment("Comment C1", "EPPlus");

            //Act
            ws.Cells["B1"].Delete(eShiftTypeDelete.Left);

            //Assert
            Assert.AreEqual("Comment A1", ws.Cells["A1"].Comment.Text);
            Assert.AreEqual("Comment C1", ws.Cells["B1"].Comment.Text);
            Assert.IsNull(ws.Cells["C1"].Comment);
        }
        [TestMethod]
        public void ValidateNameAfterDeleteShiftUp()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("InsertRangeNamesDown");
            ws.Names.Add("NameA1", ws.Cells["A1"]);
            ws.Names.Add("NameA2", ws.Cells["A2"]);
            ws.Names.Add("NameB1", ws.Cells["B1"]);
            ws.Names.Add("NameB2", ws.Cells["B2"]);
            ws.Names.Add("NameC1", ws.Cells["C1"]);
            ws.Names.Add("NameC2", ws.Cells["C2"]);

            //Act
            ws.Cells["A1"].Delete(eShiftTypeDelete.Up);

            //Assert
            Assert.AreEqual("#REF!", ws.Names["NameA1"].Address);
            Assert.AreEqual("A1", ws.Names["NameA2"].Address);
            Assert.AreEqual("B1", ws.Names["NameB1"].Address);
            Assert.AreEqual("C1", ws.Names["NameC1"].Address);
        }
        [TestMethod]
        public void ValidateNameAfterDeleteShiftUp_MustBeInsideRange()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("InsertRangeInsideNamesDown");
            ws.Names.Add("NameA2B4", ws.Cells["A2:B4"]);
            ws.Names.Add("NameB2D3", ws.Cells["B2:D3"]);
            ws.Names.Add("NameC1F3", ws.Cells["C1:F3"]);

            //Act
            ws.Cells["A2:B3"].Delete(eShiftTypeDelete.Up);

            //Assert
            Assert.AreEqual("A2:B2", ws.Names["NameA2B4"].Address);
            //Assert.AreEqual("B2:D3", ws.Names["NameB2D3"].Address);
            //Assert.AreEqual("C1:F3", ws.Names["NameC1F3"].Address);

            //ws.Cells["B2:D5"].Insert(eShiftTypeInsert.Down);
            //Assert.AreEqual("A4:B6", ws.Names["NameA2B4"].Address);
            //Assert.AreEqual("B6:D7", ws.Names["NameB2D3"].Address);
            //Assert.AreEqual("C1:F3", ws.Names["NameC1F3"].Address);

            //ws.Cells["B2:F2"].Insert(eShiftTypeInsert.Down);
            //Assert.AreEqual("A4:B6", ws.Names["NameA2B4"].Address);
            //Assert.AreEqual("B7:D8", ws.Names["NameB2D3"].Address);
            //Assert.AreEqual("C1:F4", ws.Names["NameC1F3"].Address);
        }

        [TestMethod]
        public void ValidateNamesAfterDeleteShiftLeft_MustBeInsideRange()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("InsertRangeInsideNamesRight");
            ws.Names.Add("NameB1D2", ws.Cells["B1:D2"]);
            ws.Names.Add("NameB2C4", ws.Cells["B2:D4"]);
            ws.Names.Add("NameA3C6", ws.Cells["A3:C6"]);

            //Act
            ws.Cells["B1:C2"].Insert(eShiftTypeInsert.Right);

            ////Assert
            //Assert.AreEqual("D1:F2", ws.Names["NameB1D2"].Address);
            //Assert.AreEqual("B2:D4", ws.Names["NameB2C4"].Address);
            //Assert.AreEqual("A3:C6", ws.Names["NameA3C6"].Address);

            //ws.Cells["B2:D5"].Insert(eShiftTypeInsert.Down);
            //Assert.AreEqual("D1:F2", ws.Names["NameB1D2"].Address);
            //Assert.AreEqual("B6:D8", ws.Names["NameB2C4"].Address);
            //Assert.AreEqual("A3:C6", ws.Names["NameA3C6"].Address);
        }
    }
}
