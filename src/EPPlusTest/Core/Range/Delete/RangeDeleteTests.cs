﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
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
            ws.Cells["C1"].Formula = "Sum(DeleteRow_Sheet2!C5:C10)";
            ws.Cells["D1:D2"].Formula = "Sum(DeleteRow_Sheet2!C5:C10)";

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
            Assert.AreEqual("Sum(DeleteRow_Sheet2!C5:C10)", ws.Cells["C1"].Formula);
            Assert.AreEqual("Sum(DeleteRow_Sheet2!C5:C10)", ws.Cells["D1"].Formula);

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
            Assert.AreEqual("$A$1", ws.Names["NameA2"].Address);
            Assert.AreEqual("$B$1", ws.Names["NameB1"].Address);
            Assert.AreEqual("$C$1", ws.Names["NameC1"].Address);
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
            Assert.AreEqual("$A$2:$B$2", ws.Names["NameA2B4"].Address);
            Assert.AreEqual("$B$2:$D$3", ws.Names["NameB2D3"].Address);
            Assert.AreEqual("$C$1:$F$3", ws.Names["NameC1F3"].Address);

            ws.Cells["B2:D5"].Delete(eShiftTypeDelete.Up);
            Assert.AreEqual("$A$2:$B$2", ws.Names["NameA2B4"].Address);
            Assert.AreEqual("#REF!", ws.Names["NameB2D3"].Address);
            Assert.AreEqual("$C$1:$F$3", ws.Names["NameC1F3"].Address);

            ws.Cells["B2:F2"].Delete(eShiftTypeDelete.Up);
            Assert.AreEqual("$A$2:$B$2", ws.Names["NameA2B4"].Address);
            Assert.AreEqual("#REF!", ws.Names["NameB2D3"].Address);
            Assert.AreEqual("$C$1:$F$2", ws.Names["NameC1F3"].Address);
        }

        [TestMethod]
        public void ValidateNamesAfterDeleteShiftLeft_MustBeInsideRange()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("InsertRangeInsideNamesRight");
            ws.Names.Add("NameB1D2", ws.Cells["D1:F2"]);
            ws.Names.Add("NameB2C4", ws.Cells["D2:F4"]);
            ws.Names.Add("NameA3C6", ws.Cells["A3:C6"]);

            //Act
            ws.Cells["B1:C2"].Delete(eShiftTypeDelete.Left);

            //Assert
            Assert.AreEqual("$B$1:$D$2", ws.Names["NameB1D2"].Address);
            Assert.AreEqual("$D$2:$F$4", ws.Names["NameB2C4"].Address);
            Assert.AreEqual("$A$3:$C$6", ws.Names["NameA3C6"].Address);

            ws.Cells["B2:D5"].Delete(eShiftTypeDelete.Left);
            Assert.AreEqual("$B$1:$D$2", ws.Names["NameB1D2"].Address);
            Assert.AreEqual("$B$2:$C$4", ws.Names["NameB2C4"].Address);
            Assert.AreEqual("$A$3:$C$6", ws.Names["NameA3C6"].Address);

            ws.Cells["A2:C7"].Delete(eShiftTypeDelete.Left);
            Assert.AreEqual("$B$1:$D$2", ws.Names["NameB1D2"].Address);
            Assert.AreEqual("#REF!", ws.Names["NameB2C4"].Address);
            Assert.AreEqual("#REF!", ws.Names["NameA3C6"].Address);
        }
        [TestMethod]
        public void ValidateSharedFormulasDeleteShiftUp()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DeleteRangeFormulaUp");
            ws.Cells["B1:D2"].Formula = "A1";
            ws.Cells["C3:F4"].Formula = "A1";

            //Act
            ws.Cells["B1"].Delete(eShiftTypeDelete.Up);

            //Assert
            Assert.AreEqual("A2", ws.Cells["B1"].Formula);
            Assert.AreEqual("",ws.Cells["B2"].Formula);
            Assert.AreEqual("#REF!", ws.Cells["C1"].Formula);
            Assert.AreEqual("C1", ws.Cells["D1"].Formula);
            Assert.AreEqual("A1", ws.Cells["C3"].Formula);
            Assert.AreEqual("#REF!", ws.Cells["D3"].Formula);
            Assert.AreEqual("C1", ws.Cells["E3"].Formula);
            Assert.AreEqual("D1", ws.Cells["F3"].Formula);


            Assert.AreEqual("D2", ws.Cells["F4"].Formula);
        }
        [TestMethod]
        public void ValidateSharedFormulasDeleteShiftLeft()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("DeleteRangeFormulaLeft");
            ws.Cells["B1:D2"].Formula = "A1";
            ws.Cells["C3:F4"].Formula = "A1";

            //Act
            ws.Cells["B1"].Delete(eShiftTypeDelete.Left);

            //Assert
            Assert.AreEqual("#REF!", ws.Cells["B1"].Formula);
            Assert.AreEqual("B1", ws.Cells["C1"].Formula);
            Assert.AreEqual("", ws.Cells["D1"].Formula);
            Assert.AreEqual("A2", ws.Cells["B2"].Formula);
            Assert.AreEqual("A1", ws.Cells["C3"].Formula);
            Assert.AreEqual("#REF!", ws.Cells["D3"].Formula);
            Assert.AreEqual("B1", ws.Cells["E3"].Formula);
            Assert.AreEqual("C1", ws.Cells["F3"].Formula);


            Assert.AreEqual("A1", ws.Cells["C3"].Formula);
            Assert.AreEqual("D2", ws.Cells["F4"].Formula);
        }

        [TestMethod]
        public void ValidateDeleteMergedCellsUp()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("MergedCells");
                ws.Cells["C3:E4"].Merge = true;
                ws.Cells["C2:E2"].Delete(eShiftTypeDelete.Up);

                Assert.AreEqual("C2:E3", ws.MergedCells[0]);
            }
        }
        [TestMethod]
        public void ValidateDeleteMergedCellsLeft()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("MergedCells");
                ws.Cells["C2:E3"].Merge = true;
                ws.Cells["B2:B3"].Delete(eShiftTypeDelete.Left);

                Assert.AreEqual("B2:D3", ws.MergedCells[0]);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ValidateDeleteIntoMergedCellsPartialLeftThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("MergedCells");
                ws.Cells["B2:D3"].Merge = true;
                ws.Cells["A2"].Delete(eShiftTypeDelete.Left);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ValidateDeleteIntoMergedCellsPartialUpThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("MergedCells");
                ws.Cells["B2:D3"].Merge = true;
                ws.Cells["C1"].Delete(eShiftTypeDelete.Up);
            }
        }
        [TestMethod]
        public void ValidateDeleteIntoMergedCellsPartialLeftShouldNotThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("MergedCells");
                ws.Cells["B2:D3"].Merge = true;
                ws.Cells["C1"].Delete(eShiftTypeDelete.Left);
            }
        }
        [TestMethod]
        public void ValidateDeleteIntoMergedCellsPartialUpShouldNotThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("MergedCells");
                ws.Cells["B2:D3"].Merge = true;
                ws.Cells["A2"].Delete(eShiftTypeDelete.Up);
            }
        }

        [TestMethod]
        public void ValidateDeleteMergedCellsShouldShiftUp()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("MergedCells");
                ws.Cells["B3:D4"].Merge = true;
                ws.Cells["A1:D1"].Delete(eShiftTypeDelete.Up);

                Assert.AreEqual("B2:D3", ws.MergedCells[0]);
                Assert.IsFalse(ws.Cells["B4"].Merge);
                Assert.IsFalse(ws.Cells["C4"].Merge);
                Assert.IsFalse(ws.Cells["D4"].Merge);

                Assert.IsTrue(ws.Cells["B2"].Merge);
                Assert.IsTrue(ws.Cells["C2"].Merge);
                Assert.IsTrue(ws.Cells["D2"].Merge);
                Assert.IsTrue(ws.Cells["B3"].Merge);
                Assert.IsTrue(ws.Cells["C3"].Merge);
                Assert.IsTrue(ws.Cells["D3"].Merge);

                ws.DeleteRow(1);
                Assert.AreEqual("B1:D2", ws.MergedCells[0]);
            }
        }
        [TestMethod]
        public void ValidateDeleteMergedCellsShouldBeNull()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("MergedCells");
                ws.Cells["B3:D3"].Merge = true;
                ws.Cells["B3:D3"].Delete(eShiftTypeDelete.Up);

                Assert.IsFalse(ws.Cells["B3"].Merge);
                Assert.IsFalse(ws.Cells["C3"].Merge);
                Assert.IsFalse(ws.Cells["D3"].Merge);
                Assert.IsNull(ws.MergedCells[0]);

                ws.Cells["B3:D3"].Merge = true;

                ws.DeleteRow(3);
                Assert.IsFalse(ws.Cells["B3"].Merge);
                Assert.IsFalse(ws.Cells["C3"].Merge);
                Assert.IsFalse(ws.Cells["D3"].Merge);
                Assert.IsNull(ws.MergedCells[1]);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ValidateDeleteFromTablePartialLeftThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("TableDelete");
                ws.Tables.Add(ws.Cells["B2:D3"], "table1");
                ws.Cells["A2"].Delete(eShiftTypeDelete.Left);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ValidateDeleteFromTablePartialUpThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("TableDelete");
                ws.Tables.Add(ws.Cells["B2:D3"], "table1");
                ws.Cells["C1"].Delete(eShiftTypeDelete.Up);
            }
        }
        [TestMethod]
        public void ValidateDeletFromTablePartialLeftShouldNotThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("TableDelete");
                ws.Tables.Add(ws.Cells["B2:D3"], "table1");
                ws.Cells["C1"].Delete(eShiftTypeDelete.Left);
            }
        }
        [TestMethod]
        public void ValidateDeleteFromTablePartialUpShouldNotThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("TableDelete");
                ws.Tables.Add(ws.Cells["B2:D3"], "table1");
                ws.Cells["A2"].Delete(eShiftTypeDelete.Up);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ValidateDeleteFromPivotTablePartialLeftThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("PivotTableDelete");
                ws.Cells["E5"].Value = "E5";
                ws.Cells["F5"].Value = "F5";
                ws.PivotTables.Add(ws.Cells["B2:D3"], ws.Cells["E5:F6"], "table1");
                ws.Cells["A2"].Delete(eShiftTypeDelete.Left);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ValidateDeleteFromPivotTablePartialUpThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("PivotTableDelete");
                ws.Cells["E5"].Value = "E5";
                ws.Cells["F5"].Value = "F5";
                ws.PivotTables.Add(ws.Cells["B2:D3"], ws.Cells["E5:F6"], "table1");
                ws.Cells["C1"].Delete(eShiftTypeDelete.Up);
            }
        }
        [TestMethod]
        public void ValidateDeleteFromPivotTablePartialLeftShouldNotThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("PivotTableDelte");
                ws.Cells["E5"].Value = "E5";
                ws.Cells["F5"].Value = "F5";
                ws.PivotTables.Add(ws.Cells["B2:D3"], ws.Cells["E5:F6"], "table1");
                ws.Cells["C1"].Delete(eShiftTypeDelete.Left);
            }
        }
        [TestMethod]
        public void ValidateDeleteFromPivotTablePartialUpShouldNotThrowsException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("PivotTableDelete");
                ws.Cells["E5"].Value = "E5";
                ws.Cells["F5"].Value = "F5";
                ws.PivotTables.Add(ws.Cells["B2:D3"], ws.Cells["E5:F6"], "table1");
                ws.Cells["A2"].Delete(eShiftTypeDelete.Up);
            }
        }
        [TestMethod]
        public void ValidateDeleteFromTableShouldShiftUp()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("TableDeleteShiftUp");
                var tbl = ws.Tables.Add(ws.Cells["B9:D10"], "table1");
                ws.Cells["B2:D2"].Delete(eShiftTypeDelete.Up);
                Assert.AreEqual("B8:D9", tbl.Address.Address);

                ws.Cells["A3:D3"].Delete(eShiftTypeDelete.Up);
                Assert.AreEqual("B7:D8", tbl.Address.Address);

                ws.Cells["B3:E3"].Delete(eShiftTypeDelete.Up);
                Assert.AreEqual("B6:D7", tbl.Address.Address);
            }
        }
        [TestMethod]
        public void ValidateDeleteTableShouldShiftLeft()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("TableDeleteShiftLeft");
                var tbl = ws.Tables.Add(ws.Cells["E2:F4"], "table1");
                ws.Cells["B2:B4"].Delete(eShiftTypeDelete.Left);
                Assert.AreEqual("D2:E4", tbl.Address.Address);

                ws.Cells["B1:B4"].Delete(eShiftTypeDelete.Left);
                Assert.AreEqual("C2:D4", tbl.Address.Address);

                ws.Cells["B2:B6"].Delete(eShiftTypeDelete.Left);
                Assert.AreEqual("B2:C4", tbl.Address.Address);
            }
        }
        [TestMethod]
        public void DeleteEntireTableRangeShouldDeleteTable()
        {
            using (var p = new ExcelPackage())
            {
                //Setup
                var ws = p.Workbook.Worksheets.Add("TableDeleteFull");
                var tbl = ws.Tables.Add(ws.Cells["E2:F4"], "table1");
                //Act
                ws.Cells["E2:F4"].Delete(eShiftTypeDelete.Left);
                //Assert
                Assert.AreEqual(0, ws.Tables.Count);
                Assert.IsNull(tbl.Address);
            }
        }
        [TestMethod]
        public void DeleteEntirePivotTableRangeShouldDeletePivotTable()
        {
            using (var p = new ExcelPackage())
            {
                //Setup
                var ws = p.Workbook.Worksheets.Add("PivotTableDeleteFull");
                ws.Cells["E5"].Value = "E5";
                ws.Cells["F5"].Value = "F5";
                var pt = ws.PivotTables.Add(ws.Cells["B2:D3"], ws.Cells["E5:F6"], "pivottable1");
                //Act
                ws.Cells["B2:D3"].Delete
                    (eShiftTypeDelete.Left);
                //Assert
                Assert.AreEqual(0, ws.PivotTables.Count);
                Assert.IsNull(pt.Address);
            }
        }

        [TestMethod]
        public void ValidateDeletePivotTableShouldShiftUp()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("PivotTableDeleteShiftUp");
                ws.Cells["E5"].Value = "E5";
                ws.Cells["F5"].Value = "F5";
                var pt = ws.PivotTables.Add(ws.Cells["B5:D6"], ws.Cells["E5:F6"], "pivottable1");
                ws.Cells["B2:D2"].Delete(eShiftTypeDelete.Up);
                Assert.AreEqual("B4:D5", pt.Address.Address);

                ws.Cells["A2:E2"].Delete(eShiftTypeDelete.Up);
                Assert.AreEqual("B3:D4", pt.Address.Address);

                ws.Cells["B5:D5"].Delete(eShiftTypeDelete.Up);
                Assert.AreEqual("B3:D4", pt.Address.Address);
            }
        }
        [TestMethod]
        public void ValidateDeletePivotTableShouldShiftLeft()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("PivotTableDeleteShiftLeft");
                ws.Cells["E5"].Value = "E5";
                ws.Cells["F5"].Value = "F5";
                var pt = ws.PivotTables.Add(ws.Cells["F2:G3"], ws.Cells["E5:F6"], "pivottable1");
                ws.Cells["B2:B3"].Delete(eShiftTypeDelete.Left);
                Assert.AreEqual("E2:F3", pt.Address.Address);
                ws.Cells["B1:B4"].Delete(eShiftTypeDelete.Left);
                Assert.AreEqual("D2:E3", pt.Address.Address);
                ws.Cells["F2:F3"].Delete(eShiftTypeDelete.Left);
                Assert.AreEqual("D2:E3", pt.Address.Address);
            }
        }

        #region Data validation
        [TestMethod]
        public void ValidateDatavalidationFullShiftUp()
        {
            var ws = _pck.Workbook.Worksheets.Add("DataValShiftUpFull");
            var any = ws.DataValidations.AddAnyValidation("B2:E5");

            ws.Cells["A1:E1"].Delete(eShiftTypeDelete.Up);

            Assert.AreEqual("B1:E4", any.Address.Address);
        }
        [TestMethod]
        public void ValidateDatavalidationPartialShiftUp_Left()
        {
            var ws = _pck.Workbook.Worksheets.Add("DataValPartialUpFullL");
            var any = ws.DataValidations.AddAnyValidation("B2:E5");

            ws.Cells["A1:C1"].Delete(eShiftTypeDelete.Up);

            Assert.AreEqual("B1:C4,D2:E5", any.Address.Address);
        }
        [TestMethod]
        public void ValidateDatavalidationPartialShiftUp_Inside()
        {
            var ws = _pck.Workbook.Worksheets.Add("DataValPartialUpFullI");
            var any = ws.DataValidations.AddAnyValidation("B2:E5");

            ws.Cells["C1:D1"].Delete(eShiftTypeDelete.Up);

            Assert.AreEqual("B2:B5,C1:D4,E2:E5", any.Address.Address);
        }


        [TestMethod]
        public void ValidateDatavalidationPartialShiftUp_Right()
        {
            var ws = _pck.Workbook.Worksheets.Add("DataValPartialUpFullR");
            var any = ws.DataValidations.AddAnyValidation("B2:E5");

            ws.Cells["C1:E1"].Delete(eShiftTypeDelete.Up);

            Assert.AreEqual("B2:B5,C1:E4", any.Address.Address);
        }
        [TestMethod]
        public void ValidateDatavalidationPartialShiftLeft_Top()
        {
            var ws = _pck.Workbook.Worksheets.Add("DataValPartialLeftFullTop");
            var any = ws.DataValidations.AddAnyValidation("B2:E5");

            ws.Cells["A2:A4"].Delete(eShiftTypeDelete.Left);

            Assert.AreEqual("A2:D4,B5:E5", any.Address.Address);
        }
        [TestMethod]
        public void ValidateDatavalidationPartialShiftLeft_Inside()
        {
            var ws = _pck.Workbook.Worksheets.Add("DataValPartialLeftFullIns");
            var any = ws.DataValidations.AddAnyValidation("B2:E5");

            ws.Cells["A3:A4"].Delete(eShiftTypeDelete.Left);

            Assert.AreEqual("B2:E2,A3:D4,B5:E5", any.Address.Address);
        }

        [TestMethod]
        public void ValidateDatavalPartialShiftLeft_Bottom()
        {
            var ws = _pck.Workbook.Worksheets.Add("DataValPartialLeftFullBottom");
            var any = ws.DataValidations.AddAnyValidation("B2:E5");

            ws.Cells["A3:A6"].Delete(eShiftTypeDelete.Left);

            Assert.AreEqual("B2:E2,A3:D5", any.Address.Address);
        }

        [TestMethod]
        public void ValidateDatavalidationFullShiftLeft()
        {
            var ws = _pck.Workbook.Worksheets.Add("DataValidationShiftLeftFull");
            var any = ws.DataValidations.AddAnyValidation("B2:E5");

            ws.Cells["A2:A5"].Delete(eShiftTypeDelete.Left);

            Assert.AreEqual("A2:D5", any.Address.Address);
        }
        [TestMethod]
        public void CheckDatavalidationFormulaAfterDeletingRow()
        {
            using (var p = new ExcelPackage())
            {
                // Create a worksheet with conditional formatting 
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var dv = ws.DataValidations.AddCustomValidation("B5:G5");
                dv.Formula.ExcelFormula = "=(B$4=0)";

                // Delete a row before the column being referenced by the CF formula
                ws.DeleteRow(2);

                // Check the conditional formatting formula has been updated
                dv = ws.DataValidations[0].As.CustomValidation;
                Assert.AreEqual("=(B$3=0)", dv.Formula.ExcelFormula);
            }
        }
        [TestMethod]
        public void CheckDatavalidationFormulaAfterDeletingColumn()
        {
            using (var p = new ExcelPackage())
            {
                // Create a worksheet with conditional formatting 
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var dv = ws.DataValidations.AddCustomValidation("E2:E7");
                dv.Formula.ExcelFormula = "=($D2=0)";

                // Delete a column before the column being referenced by the CF formula
                ws.DeleteColumn(2);

                // Check the conditional formatting formula has been updated
                dv = ws.DataValidations[0].As.CustomValidation;
                Assert.AreEqual("=($C2=0)", dv.Formula.ExcelFormula);
            }
        }

        #endregion
        #region Conditional formatting
        [TestMethod]
        public void ValidateConditionalFormattingFullShiftUp()
        {
            var ws = _pck.Workbook.Worksheets.Add("CondFormShiftUpFull");
            var cf = ws.ConditionalFormatting.AddAboveAverage(new ExcelAddress("B2:E5"));
            cf.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent1);
            ws.Cells["A1:E1"].Delete(eShiftTypeDelete.Up);

            Assert.AreEqual("B1:E4", cf.Address.Address);
        }
        [TestMethod]
        public void ValidateConditionalFormattingPartialShiftUp_Left()
        {
            var ws = _pck.Workbook.Worksheets.Add("CondFormPartialUpFullL");
            var cf = ws.ConditionalFormatting.AddAboveAverage(new ExcelAddress("B2:E5"));
            cf.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent1);

            ws.Cells["A2:C2"].Delete(eShiftTypeDelete.Up);

            Assert.AreEqual("B2:C4,D2:E5", cf.Address.Address);
        }
        [TestMethod]
        public void ValidateConditionalFormattingShiftUp_Inside()
        {
            var ws = _pck.Workbook.Worksheets.Add("CondFormPartialUpFullI");
            var cf = ws.ConditionalFormatting.AddAboveAverage(new ExcelAddress("B2:E5"));
            cf.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent1);

            ws.Cells["C2:D2"].Delete(eShiftTypeDelete.Up);

            Assert.AreEqual("B2:B5,C2:D4,E2:E5", cf.Address.Address);
        }


        [TestMethod]
        public void ValidateConditionalFormattingShiftUp_Right()
        {
            var ws = _pck.Workbook.Worksheets.Add("CondFormPartialUpFullR");
            var cf = ws.ConditionalFormatting.AddAboveAverage(new ExcelAddress("B2:E5"));
            cf.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent1);

            ws.Cells["C2:E3"].Delete(eShiftTypeDelete.Up);

            Assert.AreEqual("B2:B5,C2:E3", cf.Address.Address);
        }
        [TestMethod]
        public void ValidateConditionalFormattingPartialShiftLeft_Top()
        {
            var ws = _pck.Workbook.Worksheets.Add("CondFormPartialRightFullTop");
            var cf = ws.ConditionalFormatting.AddAboveAverage(new ExcelAddress("B2:E5"));
            cf.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent1);

            ws.Cells["A2:A4"].Delete(eShiftTypeDelete.Left);

            Assert.AreEqual("A2:D4,B5:E5", cf.Address.Address);
        }
        [TestMethod]
        public void ValidateConditionalFormattingPartialShiftLeft_Inside()
        {
            var ws = _pck.Workbook.Worksheets.Add("CondFormPartialRightFullIns");
            var cf = ws.ConditionalFormatting.AddAboveAverage(new ExcelAddress("B2:E5"));
            cf.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent1);

            ws.Cells["A3:A4"].Delete(eShiftTypeDelete.Left);

            Assert.AreEqual("B2:E2,A3:D4,B5:E5", cf.Address.Address);
        }

        [TestMethod]
        public void ValidateConditionalFormattingShiftLeft_Bottom()
        {
            var ws = _pck.Workbook.Worksheets.Add("CondFormPartialDownFullBottom");
            var cf = ws.ConditionalFormatting.AddAboveAverage(new ExcelAddress("B2:E5"));
            cf.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent1);

            ws.Cells["A3:A6"].Delete(eShiftTypeDelete.Left);

            Assert.AreEqual("B2:E2,A3:D5", cf.Address.Address);
        }

        [TestMethod]
        public void ValidateConditionalFormattingFullShiftLeft()
        {
            var ws = _pck.Workbook.Worksheets.Add("CondFormShiftRightFull");
            var cf = ws.ConditionalFormatting.AddAboveAverage(new ExcelAddress("B2:E5"));
            cf.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent1);

            ws.Cells["A2:A5"].Delete(eShiftTypeDelete.Left);

            Assert.AreEqual("A2:D5", cf.Address.Address);
        }
        [TestMethod]
        public void CheckConditionalFormattingFormulaAfterDeletingRow()
        {
            using (var p = new ExcelPackage())
            {
                // Create a worksheet with conditional formatting 
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var cf = ws.ConditionalFormatting.AddExpression(ws.Cells["B5:G5"]);
                cf.Formula = "=(B$4=0)";

                // Delete a row before the column being referenced by the CF formula
                ws.DeleteRow(2);

                // Check the conditional formatting formula has been updated
                cf = (IExcelConditionalFormattingExpression)ws.ConditionalFormatting[0];
                Assert.AreEqual("=(B$3=0)", cf.Formula);
            }
        }
        [TestMethod]
        public void CheckConditionalFormattingFormulaAfterDeletingColumn()
        {
            using (var p = new ExcelPackage())
            {
                // Create a worksheet with conditional formatting 
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var cf = ws.ConditionalFormatting.AddExpression(ws.Cells["E2:E7"]);
                cf.Formula = "=($D2=0)";

                // Delete a column before the column being referenced by the CF formula
                ws.DeleteColumn(2);

                // Check the conditional formatting formula has been updated
                cf = (IExcelConditionalFormattingExpression)ws.ConditionalFormatting[0];
                Assert.AreEqual("=($C2=0)", cf.Formula);
            }
        }
        #endregion

        [TestMethod]
        public void ValidateFilterShiftUp()
        {
            var ws = _pck.Workbook.Worksheets.Add("AutoFilterShiftUp");
            LoadTestdata(ws);
            ws.AutoFilterAddress = new ExcelAddressBase("A2:D100");
            ws.Cells["A1:D1"].Delete(eShiftTypeDelete.Up);
            Assert.AreEqual("A1:D99", ws.AutoFilterAddress.Address);
            ws.Cells["A50:D50"].Delete(eShiftTypeDelete.Up);
            Assert.AreEqual("A1:D98", ws.AutoFilterAddress.Address);
        }
        [TestMethod]
        public void ValidateFilterDeleteFirstRow()
        {
            var ws = _pck.Workbook.Worksheets.Add("AutoFilterDeleteFirstRow");
            LoadTestdata(ws);
            ws.AutoFilterAddress = new ExcelAddressBase("A1:D100");
            ws.Cells["A1:D1"].Delete(eShiftTypeDelete.Up);
            Assert.IsNull(ws.AutoFilterAddress);
        }
        [TestMethod]
        public void ValidateFilterShiftLeft()
        {
            var ws = _pck.Workbook.Worksheets.Add("AutoFilterShiftLeft");
            LoadTestdata(ws, 100, 2);
            ws.AutoFilterAddress = new ExcelAddressBase("B1:E100");
            ws.Cells["A1:A100"].Delete(eShiftTypeDelete.Left);
            Assert.AreEqual("A1:D100", ws.AutoFilterAddress.Address);
            ws.Cells["C1:C100"].Delete(eShiftTypeDelete.Left); 
            Assert.AreEqual("A1:C100", ws.AutoFilterAddress.Address);
        }
        [TestMethod]
        public void ValidateFilterDeleteRow()
        {
            var ws = _pck.Workbook.Worksheets.Add("AutoFilterDeleteRow");
            LoadTestdata(ws);
            ws.AutoFilterAddress = new ExcelAddressBase("A2:D100");
            ws.DeleteRow(1, 1);
            Assert.AreEqual("A1:D99", ws.AutoFilterAddress.Address);
            ws.DeleteRow(5, 2);
            Assert.AreEqual("A1:D97", ws.AutoFilterAddress.Address);
        }
        [TestMethod]
        public void ValidateFilterDeleteRowFirstRow()
        {
            var ws = _pck.Workbook.Worksheets.Add("AutoFilterDeleteRowFirstRow");
            LoadTestdata(ws);
            ws.AutoFilterAddress = new ExcelAddressBase("A1:D100");
            ws.DeleteRow(1);
            Assert.IsNull(ws.AutoFilterAddress);
        }
        [TestMethod]
        public void ValidateFilterDeleteColumn()
        {
            var ws = _pck.Workbook.Worksheets.Add("AutoFilterDeleteCol");
            LoadTestdata(ws);
            ws.AutoFilterAddress = new ExcelAddressBase("B1:E100");
            ws.DeleteColumn(1, 1);
            Assert.AreEqual("A1:D100", ws.AutoFilterAddress.Address);
            ws.DeleteColumn(1, 2);
            Assert.AreEqual("A1:B100", ws.AutoFilterAddress.Address);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ValidateFilterShiftUpPartial()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("AutoFilterShiftUpPart");
                LoadTestdata(ws);
                ws.AutoFilterAddress = new ExcelAddressBase("A1:D100");
                ws.Cells["A1:C1"].Delete(eShiftTypeDelete.Up);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ValidateFilterShiftLeftPartial()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("AutoFilterShiftLeftPart");
                LoadTestdata(ws);
                ws.AutoFilterAddress = new ExcelAddressBase("A1:D100");
                ws.Cells["A1:A99"].Delete(eShiftTypeDelete.Left);
            }
        }
        [TestMethod]
        public void ValidateSparkLineShiftLeft()
        {
            var ws = _pck.Workbook.Worksheets.Add("SparklineShiftLeft");
            LoadTestdata(ws, 10, 2);
            ws.SparklineGroups.Add(OfficeOpenXml.Sparkline.eSparklineType.Line, ws.Cells["F2:F10"], ws.Cells["B2:E10"]);
            ws.Cells["F5"].Delete(eShiftTypeDelete.Left);
            Assert.AreEqual("F6", ws.SparklineGroups[0].Sparklines[3].Cell.Address);
            ws.Cells["A1:A10"].Delete(eShiftTypeDelete.Left);
            Assert.AreEqual("A2:D10", ws.SparklineGroups[0].DataRange.Address);
            ws.Cells["B2:D2"].Delete(eShiftTypeDelete.Left);
            Assert.AreEqual("SparklineShiftLeft!A2", ws.SparklineGroups[0].Sparklines[0].RangeAddress.Address);
            ws.Cells["A3:D3"].Delete(eShiftTypeDelete.Left);
            Assert.IsNull(ws.SparklineGroups[0].Sparklines[1].RangeAddress);
        }
        [TestMethod]
        public void ValidateSparkLineShiftUp()
        {
            var ws = _pck.Workbook.Worksheets.Add("SparklineShiftUp");
            LoadTestdata(ws, 10);
            ws.SparklineGroups.Add(OfficeOpenXml.Sparkline.eSparklineType.Column, ws.Cells["F2:F10"], ws.Cells["B2:E10"]);
            ws.Cells["F5"].Delete(eShiftTypeDelete.Up);
            Assert.AreEqual("F5", ws.SparklineGroups[0].Sparklines[3].Cell.Address);
            Assert.AreEqual("SparklineShiftUp!B6:E6", ws.SparklineGroups[0].Sparklines[3].RangeAddress.Address);
            ws.Cells["A1:E1"].Delete(eShiftTypeDelete.Up);
            Assert.AreEqual("B1:E9", ws.SparklineGroups[0].DataRange.Address);
        }
        [TestMethod]
        public void ValidateSparkLineDeleteRow()
        {
            var ws = _pck.Workbook.Worksheets.Add("SparklineDeleteRow");
            LoadTestdata(ws, 10);
            ws.SparklineGroups.Add(OfficeOpenXml.Sparkline.eSparklineType.Column, ws.Cells["E2:E10"], ws.Cells["A2:D10"]);
            ws.DeleteRow(5, 1);
            Assert.AreEqual("E5", ws.SparklineGroups[0].Sparklines[3].Cell.Address);
            ws.DeleteRow(1, 1);
            Assert.AreEqual("A1:D8", ws.SparklineGroups[0].DataRange.Address);
        }
        [TestMethod]
        public void ValidateSparkLineInsertColumn()
        {
            var ws = _pck.Workbook.Worksheets.Add("SparklineDeleteColumn");
            LoadTestdata(ws, 10);
            ws.SparklineGroups.Add(OfficeOpenXml.Sparkline.eSparklineType.Column, ws.Cells["E2:E10"], ws.Cells["A2:D10"]);
            ws.DeleteColumn(2, 1);
            Assert.AreEqual("D5", ws.SparklineGroups[0].Sparklines[3].Cell.Address);
            Assert.AreEqual("A5:C5", ws.SparklineGroups[0].Sparklines[3].RangeAddress.FirstAddress);
            ws.DeleteColumn(1, 1);
            Assert.AreEqual("A2:B10", ws.SparklineGroups[0].DataRange.Address);
        }
        [TestMethod]
        public void DeleteFromTemplate1()
        {
            using (var p = OpenTemplatePackage("InsertDeleteTemplate.xlsx"))
            {
                var ws = p.Workbook.Worksheets["C3R"];

                var ws2 = ws.Workbook.Worksheets.Add("C3R-2", ws);
                ws.Cells["G49:G52"].Delete(eShiftTypeDelete.Up);
                ws2.Cells["G49:G52"].Delete(eShiftTypeDelete.Left);

                SaveWorkbook("DeleteTest1.xlsx", p);
            }
        }
        [TestMethod]
        public void DeleteFromTemplate2()
        {
            using (var p = OpenTemplatePackage("InsertDeleteTemplate.xlsx"))
            {
                var ws = p.Workbook.Worksheets["C3R"];
                var ws2 = ws.Workbook.Worksheets.Add("C3R-2", ws);
                ws.Cells["L49:L52"].Delete(eShiftTypeDelete.Up);
                ws2.Cells["L49:L52"].Delete(eShiftTypeDelete.Left);

                SaveWorkbook("DeleteTest2.xlsx", p);
            }
        }
        [TestMethod]
        public void ValidateConditionalFormattingDeleteColumnMultiRange()
        {
            var ws = _pck.Workbook.Worksheets.Add("CondFormPartialUpMR");
            var cf = ws.ConditionalFormatting.AddAboveAverage(new ExcelAddress("B2:E5,D3:E5"));
            cf.Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent1);

            ws.DeleteColumn(4);

            Assert.AreEqual("B2:D5,D3:D5", cf.Address.Address);
        }
        [TestMethod]
        public void ValidateColumnShifting()
        {
            var ws = _pck.Workbook.Worksheets.Add("ColumnDelete");
            var col1 = ws.Column(3);
            col1.Width = 3;
            var col2 = ws.Column(4);
            col2.Width = 4;
            var col3 = ws.Column(6);
            col3.Width = 6;
            col3.ColumnMax = 8;

            var col4 = ws.Column(14);
            col4.Width = 14;
            col4.ColumnMax = 18;

            ws.DeleteColumn(1, 2);
            Assert.AreEqual(3, ws.Column(1).Width);
            Assert.AreEqual(4, ws.Column(2).Width);
            Assert.AreEqual(6, ws.Column(4).Width);
            ws.DeleteColumn(2, 2);
            Assert.AreEqual(6, ws.Column(2).Width);
            Assert.AreEqual(6, ws.Column(3).Width);
            Assert.AreEqual(6, ws.Column(4).Width);
            ws.DeleteColumn(1, 2);
            Assert.AreEqual(6, ws.Column(1).Width);
            Assert.AreEqual(6, ws.Column(2).Width);
            Assert.AreEqual(9.140625, ws.Column(3).Width);
        }
        [TestMethod]
        public void TestDeleteColumnsWithConditionalFormatting()
        {
            using (var pck = new ExcelPackage())
            {
                // Add a sheet with conditional formatting over multiple ranges
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                var cf = wks.ConditionalFormatting.AddExpression(new ExcelAddress("B:C,E:F,H:I,K:L"));
                cf.Formula = "=($A$1=TRUE)";

                // Delete columns K:L
                wks.DeleteColumn(11, 2);
                Assert.AreEqual("B:C,E:F,H:I", cf.Address.Address);
                // Delete columns E:I
                wks.DeleteColumn(5, 5);

                Assert.AreEqual("B:C", cf.Address.Address);
            }
        }
        [TestMethod]
        public void ValidateDeleteColumnFixedAddresses()
        {
            using(var p=new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Names.Add("TestName1", ws.Cells["$A$1"]);
                ws.Names.Add("TestName2", ws.Cells["$B$1"]); 
                ws.Names.Add("TestName3", ws.Cells["$C$1"]); 
                ws.Names.Add("TestName4", ws.Cells["$B$3:$D$3"]);
                ws.Names.Add("TestName5", ws.Cells["$A$5:$C$5"]);
                ws.Names.Add("TestName6", ws.Cells["$B$7:$C$7"]);

                //Assert
                ws.DeleteColumn(2, 2);

                //Check that the named ranges have been deleted/modified as appropriate
                Assert.AreEqual("$A$1", ws.Names["TestName1"].LocalAddress);
                Assert.AreEqual("#REF!", ws.Names["TestName2"].LocalAddress);
                Assert.AreEqual("#REF!", ws.Names["TestName3"].LocalAddress);
                Assert.AreEqual("$B$3", ws.Names["TestName4"].LocalAddress);
                Assert.AreEqual("$A$5", ws.Names["TestName5"].LocalAddress);
                Assert.AreEqual("#REF!", ws.Names["TestName6"].LocalAddress);
            }
        }
        [TestMethod]
        public void ValidateDeleteRowFixedAddresses()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Names.Add("TestName1", ws.Cells["$A$1"]);
                ws.Names.Add("TestName2", ws.Cells["$A$2"]);
                ws.Names.Add("TestName3", ws.Cells["$A$3"]);
                ws.Names.Add("TestName4", ws.Cells["$C$2:$C$4"]);
                ws.Names.Add("TestName5", ws.Cells["$E$1:$E$3"]);
                ws.Names.Add("TestName6", ws.Cells["$G$2:$G$3"]);

                //Assert
                ws.DeleteRow(2, 2);

                //Check that the named ranges have been deleted/modified as appropriate
                Assert.AreEqual("$A$1", ws.Names["TestName1"].LocalAddress);
                Assert.AreEqual("#REF!", ws.Names["TestName2"].LocalAddress);
                Assert.AreEqual("#REF!", ws.Names["TestName3"].LocalAddress);
                Assert.AreEqual("$C$2", ws.Names["TestName4"].LocalAddress);
                Assert.AreEqual("$E$1", ws.Names["TestName5"].LocalAddress);
                Assert.AreEqual("#REF!", ws.Names["TestName6"].LocalAddress);
            }
        }
        [TestMethod]
        public void TestColumnWidthsAfterDeletingColumn()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                var col = ws.Column(3);
                col.ColumnMax = 5;
                col.Width = 18;

                col = ws.Column(7);
                col.ColumnMax = 9;
                col.Width = 19;


                // Delete column 4 & 7-8
                ws.DeleteColumn(4, 1);
                ws.DeleteColumn(7, 2);

                //Assert
                Assert.AreEqual(18, ws.Column(3).Width);
                Assert.AreEqual(18, ws.Column(4).Width);
                Assert.AreEqual(ws.DefaultColWidth, ws.Column(5).Width);

                Assert.AreEqual(19, ws.Column(6).Width);
                Assert.AreEqual(ws.DefaultColWidth, ws.Column(7).Width);
            }
        }
        [TestMethod]
        public void ValidateTableCalculatedColumnFormulasAfterDeleteRowAndDeleteColumn()
        {
            //Test created from issue #484 - https://github.com/EPPlusSoftware/EPPlus/issues/484
            var ws = _pck.Workbook.Worksheets.Add("DeleteCalculateColumnFormula");

            var tbl1 = ws.Tables.Add(ws.Cells["A11:C13"], "Table3");
            tbl1.Columns[2].CalculatedColumnFormula = "A12+B12";

            var tbl2 = ws.Tables.Add(ws.Cells["E11:G13"], "Table4");
            tbl2.Columns[2].CalculatedColumnFormula = "A12+F12";

            // Check the formulas have been set correctly
            Assert.AreEqual("A12+B12", ws.Cells["C12"].Formula);
            Assert.AreEqual("A12+F12", ws.Cells["G12"].Formula);
            Assert.AreEqual("A13+F13", ws.Cells["G13"].Formula);
            Assert.AreEqual("A12+B12", tbl1.Columns[2].CalculatedColumnFormula);
            Assert.AreEqual("A12+F12", tbl2.Columns["Column3"].CalculatedColumnFormula);

            //Delete two rows above the tables 
            ws.DeleteRow(5, 2);
            //Delete the column between the tables
            ws.DeleteColumn(4, 1);

            //Check the formulas were updated
            Assert.AreEqual("A10+B10", ws.Cells["C10"].Formula);
            Assert.AreEqual("A10+E10", ws.Cells["F10"].Formula);
            Assert.AreEqual("A10+B10", tbl1.Columns[2].CalculatedColumnFormula);
            Assert.AreEqual("A10+E10", tbl2.Columns[2].CalculatedColumnFormula);
        }
        [TestMethod]
        public void ValidateTableCalculatedColumnFormulasAfterDeleteRange()
        {
            //Test created from issue #484 - https://github.com/EPPlusSoftware/EPPlus/issues/484
            var ws = _pck.Workbook.Worksheets.Add("DeleteCalcColumnFormulaRange");

            var tbl1 = ws.Tables.Add(ws.Cells["A11:C13"], "Table1");
            tbl1.Columns[2].CalculatedColumnFormula = "A12+B12";

            var tbl2 = ws.Tables.Add(ws.Cells["E11:G13"], "Table2");
            tbl2.Columns[2].CalculatedColumnFormula = "A12+F12";

            // Check the formulas have been set correctly
            Assert.AreEqual("A12+B12", ws.Cells["C12"].Formula);
            Assert.AreEqual("A12+F12", ws.Cells["G12"].Formula);
            Assert.AreEqual("A13+F13", ws.Cells["G13"].Formula);
            Assert.AreEqual("A12+B12", tbl1.Columns[2].CalculatedColumnFormula);
            Assert.AreEqual("A12+F12", tbl2.Columns["Column3"].CalculatedColumnFormula);

            //Delete two rows above the tables 
            ws.Cells["A2:D2"].Delete(eShiftTypeDelete.Up);
            //Delete the column between the tables
            ws.Cells["D1:D20"].Delete(eShiftTypeDelete.Left);

            //Check the formulas were updated
            Assert.AreEqual("A11+B11", ws.Cells["C11"].Formula);
            Assert.AreEqual("A11+E12", ws.Cells["F12"].Formula);
            Assert.AreEqual("A11+B11", tbl1.Columns[2].CalculatedColumnFormula);
            Assert.AreEqual("A11+E12", tbl2.Columns[2].CalculatedColumnFormula);
        }
        [TestMethod]
        public void DeleteEntireColumnAndShiftFormulas()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet 1");
            sheet.Cells["A1:A5"].FillNumber(x => x.StartValue = 1);
            sheet.Cells["B1:B5"].FillNumber(x => x.StartValue = 1);
            sheet.Cells["C1:C5"].Formula = "A1:A5 + 1";
            sheet.Cells["D1:D5"].Formula = "C1:C5 + 1";
            sheet.Cells["A1:D5"].Style.Fill.SetBackground(Color.LightYellow);
            sheet.Cells["B1"].Delete(eShiftTypeDelete.EntireColumn);
            sheet.Calculate();
            Assert.AreEqual(2d, sheet.Cells["B1"].Value, "Column C was not correctly shifted to B");
            Assert.AreEqual(3d, sheet.Cells["C1"].Value, "Column D was not correctly shifted to C");
        }
        [TestMethod]
        public void DeleteEntireRowAndShiftFormulas()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet 1");
            sheet.Cells["A1:E1"].FillNumber(x => x.StartValue = 1);
            sheet.Cells["A2:E2"].FillNumber(x => x.StartValue = 1);
            sheet.Cells["A3:E3"].Formula = "A1:E1 + 1";
            sheet.Cells["A4:E4"].Formula = "A3:E3 + 1";
            sheet.Cells["A5:E5"].Style.Fill.SetBackground(Color.LightYellow);
            sheet.Cells["B2"].Delete(eShiftTypeDelete.EntireRow);
            sheet.Calculate();
            Assert.AreEqual(2d, sheet.Cells["A2"].Value, "Row 3 was not correctly shifted to 2");
            Assert.AreEqual(3d, sheet.Cells["A3"].Value, "Row 4 was not correctly shifted to 3");
        }
		[TestMethod]
		public void DeleteRowValidateArrayFormula()
		{
			using var package = new ExcelPackage();
			var sheet = package.Workbook.Worksheets.Add("Sheet 1");
			sheet.Cells["A2"].Formula = "XLOOKUP($A$3,$B:$B,$C:$C)";
            sheet.Calculate();
            sheet.DeleteRow(1);

            Assert.AreEqual("XLOOKUP($A$2,$B:$B,$C:$C)", sheet.Cells["A1"].Formula);
		}

	}
}
