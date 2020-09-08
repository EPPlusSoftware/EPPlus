using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Table.PivotTable
{
    [TestClass]
    public class PivotTableValueFilterTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("PivotTableValueFilters.xlsx", true);
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
        public void AddValueEqualFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("ValueEqual");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddValueFilter(ePivotTableValueFilterType.ValueEqual, 0, 5);
        }
        [TestMethod]
        public void AddValueNotEqualFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("ValueNotEqual");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
            pt.RowFields.Add(pt.Fields[4]);
            var df=pt.DataFields.Add(pt.Fields[3]);
            pt.Fields[4].Filters.AddValueFilter(ePivotTableValueFilterType.ValueNotEqual, 0, 12.2);
            //pt.Fields[4].Filters.AddValueFilter(ePivotTableValueFilterType.ValueNotEqual, 0, 85.2);
        }
        [TestMethod]
        public void AddValueGreaterThanFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("ValueGreaterThan");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddValueFilter(ePivotTableValueFilterType.ValueGreaterThan, 0, 12.2);
        }
        [TestMethod]
        public void AddValueGreaterThanOrEqualFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("ValueGreaterThanOrEqual");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddValueFilter(ePivotTableValueFilterType.ValueGreaterThanOrEqual, 0, 12.2);
        }
        [TestMethod]
        public void AddValueLessThanFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("ValueLessThan");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddValueFilter(ePivotTableValueFilterType.ValueLessThan, 0, 12.2);
        }
        [TestMethod]
        public void AddValueLessThanOrEqualFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("ValueLessThanOrEqual");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddValueFilter(ePivotTableValueFilterType.ValueLessThanOrEqual, 0, 12.2);
        }
        [TestMethod]
        public void AddValueBetweenFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("ValueBetweeen");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddValueFilter(ePivotTableValueFilterType.ValueBetween, 0, 4, 10);
        }
        [TestMethod]
        public void AddValueNotBetweenFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("ValueNotBetweeen");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
            pt.RowFields.Add(pt.Fields[4]);
            var df = pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddValueFilter(ePivotTableValueFilterType.ValueNotBetween, df, 4, 10);
        }
        [TestMethod]
        public void AddTop10CountFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("ValueTop15Count");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
            pt.RowFields.Add(pt.Fields[4]);
            var df = pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddTop10Filter(ePivotTableTop10FilterType.Count, df, 15);
        }
        [TestMethod]
        public void AddTop10PercentFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("ValueTop20Percent");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
            pt.RowFields.Add(pt.Fields[4]);
            var df = pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddTop10Filter(ePivotTableTop10FilterType.Percent, df, 20);
        }
        [TestMethod]
        public void AddTop10SumFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("ValueTop25Sum");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
            pt.RowFields.Add(pt.Fields[4]);
            var df = pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddTop10Filter(ePivotTableTop10FilterType.Sum, df, 25);
        }
        [TestMethod]
        public void AddBottom10CountFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("ValueBottom15Count");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
            pt.RowFields.Add(pt.Fields[4]);
            var df = pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddTop10Filter(ePivotTableTop10FilterType.Count, df, 15, false);
        }
        [TestMethod]
        public void AddBottom10PercentFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("ValueBottom20Percent");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
            pt.RowFields.Add(pt.Fields[4]);
            var df = pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddTop10Filter(ePivotTableTop10FilterType.Percent, df, 20, false);
        }
        [TestMethod]
        public void AddBottom10SumFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("ValueBottom25Sum");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
            pt.RowFields.Add(pt.Fields[4]);
            var df = pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddTop10Filter(ePivotTableTop10FilterType.Sum, df, 25, false);
            ws.Cells["B4:D4"].Merge = true;
            ws.Cells["B4"].Clear();
        }
    }
}

