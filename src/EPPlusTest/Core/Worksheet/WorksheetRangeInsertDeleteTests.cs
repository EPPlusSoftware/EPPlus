using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Core.Worksheet
{
    [TestClass]
    public class WorksheetRangeInsertTests : TestBase
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
            var ws = _pck.Workbook.Worksheets.Add("Inser2Rows_Sheet1");
            var ws2 = _pck.Workbook.Worksheets.Add("Inser2Rows_Sheet2");
            ws.Cells["A1"].Formula = "Sum(C5:C10)";
            ws.Cells["B1:B2"].Formula = "Sum(C5:C10)";
            ws2.Cells["A1"].Formula = "Sum(Inser2Rows_Sheet1!C5:C10)";
            ws2.Cells["B1:B2"].Formula = "Sum(Inser2Rows_Sheet1!C5:C10)";

            //Act
            ws.InsertRow(3, 2);

            //Assert
            Assert.AreEqual("Sum(C7:C12)", ws.Cells["A1"].Formula);
            Assert.AreEqual("Sum(C7:C12)", ws.Cells["B1"].Formula);
            Assert.AreEqual("Sum(C8:C13)", ws.Cells["B2"].Formula);

            Assert.AreEqual("Sum(Inser2Rows_Sheet1!C7:C12)", ws2.Cells["A1"].Formula);
            Assert.AreEqual("Sum(Inser2Rows_Sheet1!C7:C12)", ws2.Cells["B1"].Formula);
            Assert.AreEqual("Sum(Inser2Rows_Sheet1!C8:C13)", ws2.Cells["B2"].Formula);
        }
        [TestMethod]
        public void ValidateFormulasAfterInsertColumn()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("InserColumn_Sheet1");
            var ws2 = _pck.Workbook.Worksheets.Add("InserColumn_Sheet2");
            ws.Cells["A1"].Formula = "Sum(E1:J1)";
            ws.Cells["B1:C1"].Formula = "Sum(E1:J1)";
            ws2.Cells["A1"].Formula = "Sum(InserColumn_Sheet1!E1:J1)";
            ws2.Cells["B1:C1"].Formula = "Sum(InserColumn_Sheet1!E1:J1)";

            //Act
            ws.InsertColumn(4, 1);

            //Assert
            Assert.AreEqual("Sum(F1:K1)", ws.Cells["A1"].Formula);
            Assert.AreEqual("Sum(F1:K1)", ws.Cells["B1"].Formula);
            Assert.AreEqual("Sum(G1:L1)", ws.Cells["C1"].Formula);

            Assert.AreEqual("Sum(InserColumn_Sheet1!F1:K1)", ws2.Cells["A1"].Formula);
            Assert.AreEqual("Sum(InserColumn_Sheet1!F1:K1)", ws2.Cells["B1"].Formula);
            Assert.AreEqual("Sum(InserColumn_Sheet1!G1:L1)", ws2.Cells["C1"].Formula);
        }
        [TestMethod]
        public void ValidateFormulasAfterInsert2Columns()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("Inser2Columns_Sheet1");
            var ws2 = _pck.Workbook.Worksheets.Add("Inser2Columns_Sheet2");
            ws.Cells["A1"].Formula = "Sum(E1:J1)";
            ws.Cells["B1:C1"].Formula = "Sum(E1:J1)";
            ws2.Cells["A1"].Formula = "Sum(Inser2Columns_Sheet1!E1:J1)";
            ws2.Cells["B1:C1"].Formula = "Sum(Inser2Columns_Sheet1!E1:J1)";

            //Act
            ws.InsertColumn(4, 2);

            //Assert
            Assert.AreEqual("Sum(G1:L1)", ws.Cells["A1"].Formula);
            Assert.AreEqual("Sum(G1:L1)", ws.Cells["B1"].Formula);
            Assert.AreEqual("Sum(H1:M1)", ws.Cells["C1"].Formula);

            Assert.AreEqual("Sum(Inser2Columns_Sheet1!G1:L1)", ws2.Cells["A1"].Formula);
            Assert.AreEqual("Sum(Inser2Columns_Sheet1!G1:L1)", ws2.Cells["B1"].Formula);
            Assert.AreEqual("Sum(Inser2Columns_Sheet1!H1:M1)", ws2.Cells["C1"].Formula);
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
            var ws = _pck.Workbook.Worksheets.Add("DeleteRow2_Sheet1");
            ws.Cells["B3:B6"].Formula = "A1";
            
            //Act
            ws.DeleteRow(2, 2);

            //Assert
            Assert.AreEqual("", ws.Cells["B1"].Formula);
            Assert.AreEqual("#REF!", ws.Cells["B2"].Formula);
            Assert.AreEqual("#REF!", ws.Cells["B3"].Formula);
            Assert.AreEqual("A2", ws.Cells["B4"].Formula);
            Assert.AreEqual("", ws.Cells["B6"].Formula);
        }

    }
}
