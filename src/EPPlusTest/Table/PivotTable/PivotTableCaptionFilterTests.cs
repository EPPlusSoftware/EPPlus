using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
    [TestClass]
    public class PivotTableCaptionFilterTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("PivotTableFilters.xlsx", true);
            var ws = _pck.Workbook.Worksheets.Add("Data1");
            var r = LoadItemData(ws);
            ws.Tables.Add(r, "Table1");
            ws = _pck.Workbook.Worksheets.Add("Data2");
            r = LoadItemData(ws);
            ws.Tables.Add(r, "Table2");
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void AddCaptionEqualFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("CaptionEqual");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
            pt.RowFields.Add(pt.Fields[1]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionEqual, "Hardware");
        }
        [TestMethod]
        public void AddCaptionNotEqualFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("CaptionNotEqual");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
            pt.RowFields.Add(pt.Fields[1]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionNotEqual, "Hardware");
        }
        [TestMethod]
        public void AddCaptionBeginsWithFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("CaptionBeginsWith");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
            pt.RowFields.Add(pt.Fields[1]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionBeginsWith, "H");
        }
        [TestMethod]
        public void AddCaptionNotBeginsWithFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("CaptionNotBeginsWith");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
            pt.RowFields.Add(pt.Fields[1]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionNotBeginsWith, "H");
        }

        [TestMethod]
        public void AddCaptionEndsWithFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("CaptionEndsWith");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
            pt.RowFields.Add(pt.Fields[1]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionEndsWith, "ware");
        }
        [TestMethod]
        public void AddCaptionNotEndsWithFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("CaptionNotEndsWith");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
            pt.RowFields.Add(pt.Fields[1]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionNotEndsWith, "ware");
        }
        [TestMethod]
        public void AddCaptionContainsFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("CaptionContains");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
            pt.RowFields.Add(pt.Fields[1]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionContains, "roc");
        }
        [TestMethod]
        public void AddCaptionNotContainsFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("CaptionNotContains");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
            pt.RowFields.Add(pt.Fields[1]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionNotContains, "roc");
        }
        [TestMethod]
        public void AddCaptionGreaterThanFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("CaptionGreaterThan");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
            pt.RowFields.Add(pt.Fields[1]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionGreaterThan, "H");
        }
        [TestMethod]
        public void AddCaptionGreaterThanOrEqualFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("CaptionGreaterThanOrEqual");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
            pt.RowFields.Add(pt.Fields[1]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionGreaterThanOrEqual, "H");
        }
        [TestMethod]
        public void AddCaptionLessThanFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("CaptionLessThan");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
            pt.RowFields.Add(pt.Fields[1]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionLessThan, "H");
        }
        [TestMethod]
        public void AddCaptionLessThanOrEqualFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("CaptionLessThanOrEqual");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
            pt.RowFields.Add(pt.Fields[1]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionLessThanOrEqual, "H");
        }

        [TestMethod]
        public void AddCaptionBetweenWithFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("CaptionBetween");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
            pt.RowFields.Add(pt.Fields[1]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionBetween, "H", "I");

            Assert.AreEqual("H", pt.Fields[1].Filters[0].StringValue1);
            Assert.AreEqual("I", pt.Fields[1].Filters[0].StringValue2);
            Assert.AreEqual(2, ((ExcelCustomFilterColumn)pt.Fields[1].Filters[0].Filter).Filters.Count);
        }
        [TestMethod]
        public void AddCaptionNotBetweenWithFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("CaptionNotBetween");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
            pt.RowFields.Add(pt.Fields[1]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[1].Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionNotBetween, "H", "I");
            Assert.AreEqual("H", pt.Fields[1].Filters[0].StringValue1);
            Assert.AreEqual("I", pt.Fields[1].Filters[0].StringValue2);
            Assert.AreEqual(2, ((ExcelCustomFilterColumn)pt.Fields[1].Filters[0].Filter).Filters.Count);
        }
    }
}
