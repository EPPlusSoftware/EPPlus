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

        [TestMethod]
        public void FilterShouldHandleNAAsIfEmptyValue()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "Joe";
                s.Cells["A2"].Value = "Anna";
                s.Cells["B1"].Value = 1;
                s.Cells["B2"].Value = 2;
                s.Cells["C1"].Formula = "FILTER(A1:A2, B1:B2 =1, NA())";
                s.Calculate();
                Assert.AreEqual("Joe", s.Cells["C1"].Value);
            }
        }

        [TestMethod]
        public void Filter_SingleColumnInclude_BroadcastsAcrossAllValueColumns()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");

                s.Cells["A1"].Value = "GL-40010 - Office Supplies";
                s.Cells["B1"].Value = 4520.75d;
                s.Cells["C1"].Value = "Not Posted";

                s.Cells["E1"].Formula = "FILTER(A1:B1, C1 <> \"Posted!\", 0)";
                s.Calculate();

                Assert.AreEqual("GL-40010 - Office Supplies", s.Cells["E1"].Value,
                    "Label-kolumnen ska behållas.");
                Assert.AreEqual(4520.75d, s.Cells["F1"].Value,
                    "Beloppskolumnen ska också behållas via broadcast.");
            }
        }
    }
}
