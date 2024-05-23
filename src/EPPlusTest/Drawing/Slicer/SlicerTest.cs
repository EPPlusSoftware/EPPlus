using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Table.PivotTable;
using System.IO;

namespace EPPlusTest.Drawing.Slicer
{
    [TestClass]
    public class SlicerTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("Slicer.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;

            SaveAndCleanup(_pck);

            if (File.Exists(fileName)) File.Copy(fileName, dirName + "\\SlicerRead.xlsx", true);
        }
        [TestMethod]
        public void AddTableSlicerDate()
        {
            var ws = _pck.Workbook.Worksheets.Add("TableSlicerDate");

            LoadTestdata(ws);
            var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table1");
            var slicer = ws.Drawings.AddTableSlicer(tbl.Columns[0]);

            slicer.FilterValues.Add(new ExcelFilterDateGroupItem(2019, 11, 4));
            slicer.FilterValues.Add(new ExcelFilterDateGroupItem(2019, 11, 5));
            slicer.FilterValues.Add(new ExcelFilterDateGroupItem(2019, 11, 7));
            slicer.FilterValues.Add(new ExcelFilterDateGroupItem(2019, 12));
            slicer.Cache.HideItemsWithNoData = true;
            slicer.SetPosition(1, 0, 5, 0);
            slicer.SetSize(200, 600);

            Assert.AreEqual(eCrossFilter.None, slicer.Cache.CrossFilter);
            Assert.IsTrue(slicer.Cache.HideItemsWithNoData);
            slicer.Cache.HideItemsWithNoData = false;       //Validate element is removed
            Assert.IsFalse(slicer.Cache.HideItemsWithNoData);
            slicer.Cache.HideItemsWithNoData = true;       //Add element again

            var slicer2 = ws.Drawings.AddTableSlicer(tbl.Columns[2]);
            slicer2.Style = eSlicerStyle.Light4;
            slicer2.LockedPosition = true;
            slicer2.ColumnCount = 3;
            slicer2.Cache.CrossFilter = eCrossFilter.None;
            slicer2.Cache.SortOrder = eSortOrder.Descending;
            slicer2.Cache.CustomListSort = false;

            slicer2.SetPosition(1, 0, 9, 0);
            slicer2.SetSize(200, 600);

            Assert.AreEqual(eCrossFilter.None, slicer2.Cache.CrossFilter);
            Assert.AreEqual(eSortOrder.Descending, slicer2.Cache.SortOrder);
            Assert.IsFalse(slicer2.Cache.CustomListSort);

            Assert.IsTrue(slicer2.LockedPosition);
            Assert.AreEqual(3, slicer2.ColumnCount);
            Assert.AreEqual("SlicerStyleLight4", slicer2.StyleName);
            Assert.AreEqual(eSlicerStyle.Light4, slicer2.Style);
        }
        [TestMethod]
        public void AddTableSlicerString()
        {
            var ws = _pck.Workbook.Worksheets.Add("TableSlicerNumber");

            LoadTestdata(ws);
            var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table2");
            var slicer = ws.Drawings.AddTableSlicer(tbl.Columns[1]);

            slicer.Style = eSlicerStyle.Dark1;
            slicer.FilterValues.Add("52");
            slicer.FilterValues.Add("53");
            slicer.FilterValues.Add("61");
            slicer.FilterValues.Add("102");
            slicer.StartItem = 50;
            slicer.ShowCaption = false;
            slicer.SetPosition(1, 0, 5, 0);
            slicer.SetSize(200, 600);

            Assert.AreEqual(50, slicer.StartItem);
            Assert.AreEqual(eSlicerStyle.Dark1, slicer.Style);
            Assert.IsFalse(slicer.ShowCaption);
        }
        [TestMethod]
        public void AddPivotTableSlicer()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotTableSlicer");

            LoadTestdata(ws);
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTable1");
            var rf = tbl.RowFields.Add(tbl.Fields[1]);
            rf.Items.Refresh();
            rf.Items[0].Hidden = true;
            var df = tbl.DataFields.Add(tbl.Fields[3]);
            df.Function = OfficeOpenXml.Table.PivotTable.DataFieldFunctions.Sum;

            var slicer = ws.Drawings.AddPivotTableSlicer(tbl.Fields[1]);
            slicer.Cache.Data.Items[0].Hidden = true;
            slicer.Cache.Data.Items[2].Hidden = true;
            slicer.Cache.Data.Items[4].Hidden = true;
            slicer.Style = eSlicerStyle.Light5;
            slicer.SetPosition(1, 0, 10, 0);
            slicer.SetSize(200, 600);
        }
        [TestMethod]
        public void AddPivotTableSlicerToTwoPivotTables()
        {
            var ws = _pck.Workbook.Worksheets.Add("SlicerPivotSameCache");
            LoadTestdata(ws);
            var p1 = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "Pivot1");
            p1.RowFields.Add(p1.Fields[0]);

            p1.DataFields.Add(p1.Fields[3]);
            var p2 = ws.PivotTables.Add(ws.Cells["K1"], p1.CacheDefinition, "Pivot2");
            p2.DataFields.Add(p2.Fields[1]);
            p2.RowFields.Add(p2.Fields[3]);
            //p2.Fields[4].AddDateGrouping(eDateGroupBy.Years | eDateGroupBy.Months | eDateGroupBy.Days);

            var slicer = ws.Drawings.AddPivotTableSlicer(p1.Fields[0]);
            slicer.Cache.PivotTables.Add(p2);
            p2.CacheDefinition.Refresh();
            slicer.Cache.Data.Items[0].Hidden = true;
            slicer.Cache.Data.Items[1].Hidden = true;
            slicer.Cache.Data.SortOrder = eSortOrder.Descending;
            slicer.Style = eSlicerStyle.Light5;
            slicer.SetPosition(1, 0, 15, 0);
            slicer.SetSize(200, 600);

            Assert.AreEqual(slicer.Cache.Data.SortOrder, eSortOrder.Descending);
            Assert.AreEqual(slicer.Style, eSlicerStyle.Light5);
            Assert.IsTrue(slicer.Cache.Data.Items[0].Hidden);
            Assert.IsTrue(slicer.Cache.Data.Items[1].Hidden);

            Assert.AreEqual(100, p1.Fields[0].Items.Count);
            Assert.IsTrue(p1.Fields[0].Items[0].Hidden);
            Assert.IsTrue(p1.Fields[0].Items[1].Hidden);

            Assert.AreEqual(100, p2.Fields[0].Items.Count);
            Assert.IsTrue(p2.Fields[0].Items[0].Hidden);
            Assert.IsTrue(p2.Fields[0].Items[1].Hidden);
        }
        [TestMethod]
        public void AddPivotTableSlicerToTwoPivotTablesWithDateGrouping()
        {
            var ws = _pck.Workbook.Worksheets.Add("SlicerPivotSameCacheDateGroup");
            LoadTestdata(ws);
            var p1 = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "Pivot1");
            p1.RowFields.Add(p1.Fields[0]);

            p1.DataFields.Add(p1.Fields[3]);
            var p2 = ws.PivotTables.Add(ws.Cells["K1"], p1.CacheDefinition, "Pivot2");
            p2.DataFields.Add(p2.Fields[1]);
            p2.RowFields.Add(p2.Fields[3]);

            p1.Fields[0].AddDateGrouping(eDateGroupBy.Years | eDateGroupBy.Months | eDateGroupBy.Days);
            p1.Fields[0].Name = "Days";
            var slicer = ws.Drawings.AddPivotTableSlicer(p1.Fields[0]);
            slicer.Cache.PivotTables.Add(p2);
            p2.CacheDefinition.Refresh();
            slicer.Cache.Data.Items[0].Hidden = true;
            slicer.Cache.Data.Items[1].Hidden = true;
            slicer.Cache.Data.SortOrder = eSortOrder.Ascending;
            slicer.Style = eSlicerStyle.Light4;
            slicer.SetPosition(1, 0, 15, 0);
            slicer.SetSize(200, 600);

            Assert.AreEqual(slicer.Cache.Data.SortOrder, eSortOrder.Ascending);
            Assert.AreEqual(slicer.Style, eSlicerStyle.Light4);
            Assert.IsTrue(slicer.Cache.Data.Items[0].Hidden);
            Assert.IsTrue(slicer.Cache.Data.Items[1].Hidden);

            Assert.AreEqual(369, p1.Fields[0].Items.Count);
            Assert.IsTrue(p1.Fields[0].Items[0].Hidden);
            Assert.IsTrue(p1.Fields[0].Items[1].Hidden);

            Assert.AreEqual(369, p2.Fields[0].Items.Count);
            Assert.IsTrue(p2.Fields[0].Items[0].Hidden);
            Assert.IsTrue(p2.Fields[0].Items[1].Hidden);
        }
        [TestMethod]
        public void AddPivotTableSlicerToTwoPivotTablesWithNumberGrouping()
        {
            var ws = _pck.Workbook.Worksheets.Add("SlicerPivotSameCacheNumberGroup");
            LoadTestdata(ws);
            var p1 = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "Pivot1");
            p1.RowFields.Add(p1.Fields[1]);

            p1.DataFields.Add(p1.Fields[3]);
            var p2 = ws.PivotTables.Add(ws.Cells["K1"], p1.CacheDefinition, "Pivot2");
            p2.DataFields.Add(p2.Fields[1]);
            p2.RowFields.Add(p2.Fields[3]);

            p1.Fields[1].AddNumericGrouping(0, 100, 5);
            var slicer = ws.Drawings.AddPivotTableSlicer(p1.Fields[1]);
            slicer.Cache.PivotTables.Add(p2);
            p2.CacheDefinition.Refresh();
            slicer.Cache.Data.Items[0].Hidden = true;
            slicer.Cache.Data.Items[1].Hidden = true;
            slicer.Cache.Data.SortOrder = eSortOrder.Descending;
            slicer.Style = eSlicerStyle.Light5;
            slicer.SetPosition(1, 0, 15, 0);
            slicer.SetSize(200, 600);
        }

        [TestMethod]
        public void RemovePivotTableSlicerIfLast()
        {
            var ws = _pck.Workbook.Worksheets.Add("RemovedPivotTableSlicerLast");

            LoadTestdata(ws);
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTable1");
            var rf = tbl.RowFields.Add(tbl.Fields[1]);
            rf.Items.Refresh();
            rf.Items[0].Hidden = true;
            var df = tbl.DataFields.Add(tbl.Fields[3]);
            df.Function = OfficeOpenXml.Table.PivotTable.DataFieldFunctions.Sum;

            var slicer = ws.Drawings.AddPivotTableSlicer(tbl.Fields[1]);
            Assert.AreEqual(slicer, tbl.Fields[1].Slicer);

            ws.Drawings.Remove(slicer);
            Assert.IsNull(tbl.Fields[1].Slicer);
        }
        [TestMethod]
        public void RemoveOnePivotTableSlicerNotLast()
        {
            var ws = _pck.Workbook.Worksheets.Add("RemovedPivotTableSlicerNotLast");

            LoadTestdata(ws);
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTable1");
            var rf = tbl.RowFields.Add(tbl.Fields[1]);
            rf.Items.Refresh();
            rf.Items[0].Hidden = true;
            var df = tbl.DataFields.Add(tbl.Fields[3]);
            df.Function = OfficeOpenXml.Table.PivotTable.DataFieldFunctions.Sum;

            var slicer1 = ws.Drawings.AddPivotTableSlicer(tbl.Fields[1]);
            var slicer2 = ws.Drawings.AddPivotTableSlicer(tbl.Fields[2]);
            Assert.AreEqual(slicer1, tbl.Fields[1].Slicer);
            Assert.AreEqual(slicer2, tbl.Fields[2].Slicer);

            ws.Drawings.Remove(slicer1);
            Assert.IsNull(tbl.Fields[1].Slicer);
            Assert.IsNotNull(tbl.Fields[2].Slicer);
        }
        [TestMethod]
        public void RemoveTableSlicerStringIfLast()
        {
            var ws = _pck.Workbook.Worksheets.Add("TableSlicerRemoveLast");

            LoadTestdata(ws);
            var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table3");
            var slicer = ws.Drawings.AddTableSlicer(tbl.Columns[1]);

            ws.Drawings.Remove(slicer);
            Assert.IsNull(tbl.Columns[1].Slicer);
        }
        [TestMethod]
        public void RemoveTableSlicerStringIfNotLast()
        {
            var ws = _pck.Workbook.Worksheets.Add("TableSlicerRemoveNotLast");

            LoadTestdata(ws);
            var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table4");
            var slicer1 = ws.Drawings.AddTableSlicer(tbl.Columns[1]);
            var slicer2 = ws.Drawings.AddTableSlicer(tbl.Columns[2]);

            ws.Drawings.Remove(slicer1);
            Assert.IsNull(tbl.Columns[1].Slicer);
            Assert.IsNotNull(tbl.Columns[2].Slicer);
        }
        [TestMethod]
        public void AddTwoTableSlicersToSameColumn()
        {
            var ws = _pck.Workbook.Worksheets.Add("TableSlicerTwoOnSameColumn");

            LoadTestdata(ws);
            var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table5");
            var slicer1 = ws.Drawings.AddTableSlicer(tbl.Columns[2]);
            var slicer2 = ws.Drawings.AddTableSlicer(tbl.Columns[2]);
            slicer1.SetPosition(0, 0, 5, 0);
            slicer2.SetPosition(11, 0, 5, 0);
            slicer1.FilterValues.Add("Value 12");
            slicer1.FilterValues.Add("value 10");
            slicer2.FilterValues.Add("value 15");
            slicer2.StartItem = 2;
            Assert.AreEqual(2, ws.Drawings.Count);
            Assert.IsNull(tbl.Columns[1].Slicer);
            Assert.AreEqual(slicer1, tbl.Columns[2].Slicer);
        }
        [TestMethod]
        public void AddTwoPivotTableSlicersToSameField()
        {
            var ws = _pck.Workbook.Worksheets.Add("PTSlicerTwoOnSameField");

            LoadTestdata(ws);
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTable1");
            var rf = tbl.RowFields.Add(tbl.Fields[1]);
            rf.Items.Refresh();
            rf.Items[0].Hidden = true;
            var df = tbl.DataFields.Add(tbl.Fields[3]);
            df.Function = DataFieldFunctions.Sum;

            var slicer1 = ws.Drawings.AddPivotTableSlicer(tbl.Fields[2]);
            var slicer2 = ws.Drawings.AddPivotTableSlicer(tbl.Fields[2]);
            slicer1.SetPosition(0, 0, 5, 0);
            slicer2.SetPosition(11, 0, 5, 0);
            Assert.AreEqual(slicer1, tbl.Fields[2].Slicer);
        }
    }
}