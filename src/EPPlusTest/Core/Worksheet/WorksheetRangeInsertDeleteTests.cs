using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EPPlusTest.Core.Worksheet
{
    [TestClass]
    public class WorksheetRangeInsertDeleteTests : TestBase
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
            ws.Cells["A1"].Formula="Sum(C5:C10)";
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
            var wsError = _pck.Workbook.Worksheets["InsertRow_Sheet1"];
            if(wsError!=null)
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
            var wsError = _pck.Workbook.Worksheets["InsertRow_Sheet1"];
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
            var wsError = _pck.Workbook.Worksheets["InsertRow_Sheet1"];
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
            var wsError = _pck.Workbook.Worksheets["InsertRow_Sheet1"];
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
    }
}
