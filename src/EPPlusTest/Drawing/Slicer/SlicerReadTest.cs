using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Table.PivotTable;
using System.IO;

namespace EPPlusTest.Drawing.Slicer
{
    [TestClass]
    public class SlicerReadTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("SlicerRead.xlsx");
        }
        [ClassCleanup]
        public static void Cleanup()
        {
        }
        [TestMethod]
        public void ReadTableSlicerDate()
        {
            var ws = TryGetWorksheet(_pck, "TableSlicerDate");

            var slicer = ws.Tables[0].Columns[0].Slicer;
            Assert.AreEqual(4, slicer.FilterValues.Count);
            Assert.AreEqual(4, ((ExcelFilterDateGroupItem)slicer.FilterValues[0]).Day);
            Assert.AreEqual(11, ((ExcelFilterDateGroupItem)slicer.FilterValues[0]).Month);
            Assert.AreEqual(2019, ((ExcelFilterDateGroupItem)slicer.FilterValues[0]).Year);
            Assert.AreEqual(5, ((ExcelFilterDateGroupItem)slicer.FilterValues[1]).Day);
            Assert.AreEqual(7, ((ExcelFilterDateGroupItem)slicer.FilterValues[2]).Day);

            Assert.IsTrue(slicer.Cache.HideItemsWithNoData);

            Assert.AreEqual(eCrossFilter.None, slicer.Cache.CrossFilter);
            Assert.IsTrue(slicer.Cache.HideItemsWithNoData);
            slicer.Cache.HideItemsWithNoData = false;       //Validate element is removed
            Assert.IsFalse(slicer.Cache.HideItemsWithNoData);

            var slicer2 = ws.Tables[0].Columns[2].Slicer;
            Assert.AreEqual(eSlicerStyle.Light4, slicer2.Style);
            Assert.IsTrue(slicer2.LockedPosition);
            Assert.AreEqual(3, slicer2.ColumnCount);
            Assert.AreEqual(eCrossFilter.None, slicer2.Cache.CrossFilter);
            Assert.AreEqual(eSortOrder.Descending, slicer2.Cache.SortOrder);
            Assert.IsFalse(slicer2.Cache.CustomListSort);

            Assert.AreEqual("SlicerStyleLight4", slicer2.StyleName);
            Assert.AreEqual(eSlicerStyle.Light4, slicer2.Style);
        }
        [TestMethod]
        public void ReadTableSlicerString()
        {
            var ws = TryGetWorksheet(_pck, "TableSlicerNumber");

            var tbl = ws.Tables["Table2"];
            var slicer = tbl.Columns[1].Slicer;

            Assert.AreEqual(eSlicerStyle.Dark1, slicer.Style);
            Assert.AreEqual(4, slicer.FilterValues.Count);
            Assert.AreEqual("52", ((ExcelFilterValueItem)slicer.FilterValues[0]).Value);
            Assert.AreEqual("53", ((ExcelFilterValueItem)slicer.FilterValues[1]).Value);
            Assert.AreEqual("61", ((ExcelFilterValueItem)slicer.FilterValues[2]).Value);
            Assert.AreEqual("102", ((ExcelFilterValueItem)slicer.FilterValues[3]).Value);

            Assert.AreEqual(50, slicer.StartItem);
            Assert.IsFalse(slicer.ShowCaption);
        }
        [TestMethod]
        public void ReadPivotTableSlicer()
        {
            var ws = TryGetWorksheet(_pck, "PivotTableSlicer");

            var tbl = ws.PivotTables["PivotTable1"];
            var rf = tbl.RowFields[0];
            Assert.IsTrue(rf.Items[0].Hidden);
            var df=tbl.DataFields[0];
            Assert.AreEqual(DataFieldFunctions.Sum, df.Function);

            var slicer = tbl.Fields[1].Cache.Slicer;
            
            Assert.IsTrue(slicer.Cache.Data.Items[0].Hidden);
            Assert.IsTrue(slicer.Cache.Data.Items[2].Hidden);
            Assert.IsTrue(slicer.Cache.Data.Items[4].Hidden);
            Assert.AreEqual(eSlicerStyle.Light5, slicer.Style);
        }
        [TestMethod]
        public void ReadPivotTableSlicerToTwoPivotTables()
        {
            var ws = TryGetWorksheet(_pck, "SlicerPivotSameCache");
            var p1 = ws.PivotTables["Pivot1"];
            Assert.AreEqual(1, p1.RowFields.Count);
            Assert.AreEqual(1, p1.DataFields.Count);
            var p2 = ws.PivotTables["Pivot2"];

            Assert.AreEqual(1, p1.RowFields.Count);
            Assert.AreEqual(1, p1.DataFields.Count);            
            Assert.IsNotNull(p1.Fields[0].Cache.Slicer);
            Assert.AreEqual(99, p1.Fields[0].Cache.Slicer.Cache.Data.Items.Count);

            var slicer = p1.Fields[0].Cache.Slicer;
            Assert.IsTrue(slicer.Cache.Data.Items[0].Hidden);
            Assert.IsTrue(slicer.Cache.Data.Items[1].Hidden);

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
        public void ReadPivotTableSlicerToTwoPivotTablesWithDateGrouping()
        {
            var ws = TryGetWorksheet(_pck, "SlicerPivotSameCacheDateGroup");
            var p1 = ws.PivotTables["Pivot1"];
            Assert.AreEqual(3, p1.RowFields.Count);
            Assert.AreEqual(1, p1.DataFields.Count);

            var p2 = ws.PivotTables["Pivot2"];
            Assert.AreEqual(1, p2.RowFields.Count);
            Assert.AreEqual(1, p2.DataFields.Count);

            Assert.AreEqual("Days", p1.Fields[0].Name);
            var slicer = p1.Fields[0].Cache.Slicer;
            Assert.AreEqual(p2.Fields[0].Cache.Slicer, slicer);

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
        public void ReadPivotTableSlicerToTwoPivotTablesWithNumberGrouping()
        {
            var ws = TryGetWorksheet(_pck, "SlicerPivotSameCacheNumberGroup");
            var p1 = ws.PivotTables["Pivot1"];
            Assert.AreEqual(1, p1.RowFields.Count);
            Assert.AreEqual(1, p1.DataFields.Count);

            var p2 = ws.PivotTables["Pivot2"];
            Assert.AreEqual(1, p2.RowFields.Count);
            Assert.AreEqual(1, p2.DataFields.Count);

            Assert.IsInstanceOfType(p1.Fields[1].Grouping, typeof(ExcelPivotTableFieldNumericGroup));
            var slicer = p1.Fields[1].Cache.Slicer;
            Assert.AreEqual(p2.Fields[1].Cache.Slicer, slicer);
            Assert.IsTrue(slicer.Cache.Data.Items[0].Hidden);
            Assert.IsTrue(slicer.Cache.Data.Items[1].Hidden);
            Assert.AreEqual(eSortOrder.Descending, slicer.Cache.Data.SortOrder);
            Assert.AreEqual(eSlicerStyle.Light5, slicer.Style);
        }
    }
}
