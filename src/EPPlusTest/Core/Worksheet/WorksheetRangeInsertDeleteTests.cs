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
            var ws = _pck.Workbook.Worksheets.Add("Sheet1");
            var ws2 = _pck.Workbook.Worksheets.Add("Sheet2");
            ws.Cells["A1"].Formula="Sum(C5:C10)";
            ws.Cells["B1:B2"].Formula = "Sum(C5:C10)";
            ws2.Cells["A1"].Formula = "Sheet2!Sum(C5:C10)";
            ws2.Cells["B1:B2"].Formula = "Sheet2!Sum(C5:C10)";

            //Act
            ws.InsertRow(3, 1);

            //Assert
            Assert.AreEqual("Sum(C6:C11)", ws.Cells["A1"].Formula);
            Assert.AreEqual("Sum(C6:C11)", ws.Cells["B1"].Formula);
            Assert.AreEqual("Sum(C7:C12)", ws.Cells["B2"].Formula);
        }
        [TestMethod]
        public void ValidateFormulasAfterDeleteRow()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("Sheet1");
            var ws2 = _pck.Workbook.Worksheets.Add("Sheet2");
            ws.Cells["A1"].Formula = "Sum(C5:C10)";
            ws.Cells["B1:B2"].Formula = "Sum(C5:C10)";
            ws2.Cells["A1"].Formula = "Sheet2!Sum(C5:C10)";
            ws2.Cells["B1:B2"].Formula = "Sheet2!Sum(C5:C10)";

            //Act
            ws.DeleteRow(3, 1);

            //Assert
            Assert.AreEqual("Sum(C4:C9)", ws.Cells["A1"].Formula);
            Assert.AreEqual("Sum(C4:C9)", ws.Cells["B1"].Formula);
            Assert.AreEqual("Sum(C5:C10)", ws.Cells["B2"].Formula);
        }
        [TestMethod]
        public void ValidateFormulasAfterDelete2Rows()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["B3:B6"].Formula = "A1";
            //Act
            ws.DeleteRow(2, 2);

            //Assert
            Assert.AreEqual("",ws.Cells["B1"].Formula);
            Assert.AreEqual("#REF!", ws.Cells["B2"].Formula);
            Assert.AreEqual("#REF!", ws.Cells["B3"].Formula);
            Assert.AreEqual("A2", ws.Cells["B4"].Formula);
            Assert.IsNull(ws.Cells["B6"].Formula);
        }
    }
}
