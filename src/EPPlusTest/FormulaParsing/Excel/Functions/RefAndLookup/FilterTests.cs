using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class FilterTests
    {
        ExcelPackage _package;
        ExcelWorksheet _ws;

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _ws = _package.Workbook.Worksheets.Add("Sheet1");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _ws = null;
            _package.Dispose();
        }

        [TestMethod]
        public void ShouldFilterOnRow()
        {
            _ws.Cells["A1"].Value = 1;
            _ws.Cells["B1"].Value = 2;
            _ws.Cells["C1"].Value = 3;
            _ws.Cells["A2"].Value = 4;
            _ws.Cells["B2"].Value = 5;
            _ws.Cells["C2"].Value = 6;
            _ws.Cells["A3"].Value = 1;
            _ws.Cells["B3"].Value = 5;
            _ws.Cells["C3"].Value = 7;

            _ws.Cells["D1"].Formula = "FILTER(A1:C3,A1:A3=1)";
            _ws.Calculate();
            Assert.AreEqual(1, _ws.Cells["D1"].Value);
            Assert.AreEqual(2, _ws.Cells["E1"].Value);
            Assert.AreEqual(1, _ws.Cells["D2"].Value);
            Assert.AreEqual(5, _ws.Cells["E2"].Value);
        }

        [TestMethod]
        public void ShouldFilterOnColumn()
        {
            _ws.Cells["A1"].Value = 1;
            _ws.Cells["B1"].Value = 2;
            _ws.Cells["C1"].Value = 3;
            _ws.Cells["A2"].Value = 4;
            _ws.Cells["B2"].Value = 5;
            _ws.Cells["C2"].Value = 6;
            _ws.Cells["A3"].Value = 1;
            _ws.Cells["B3"].Value = 5;
            _ws.Cells["C3"].Value = 7;

            _ws.Cells["D1"].Formula = "FILTER(A1:C3,A1:C1=1)";
            _ws.Calculate();
            Assert.AreEqual(1, _ws.Cells["D1"].Value);
            Assert.AreEqual(4, _ws.Cells["D2"].Value);
            Assert.AreEqual(1, _ws.Cells["D3"].Value);
        }
    }
}
