using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Table.PivotTable
{
    [TestClass]
    public class PivotTableCacheTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("PivotCacheTable.xlsx", true);
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
        public void ValidateSameCache()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotSameCache");
            var p1 = ws.PivotTables.Add(ws.Cells["A1"], _pck.Workbook.Worksheets[0].Tables[0].Range, "Pivot1");
            p1.RowFields.Add(p1.Fields[0]);
            p1.DataFields.Add(p1.Fields[3]);
            var p2 = ws.PivotTables.Add(ws.Cells["K1"], p1.CacheDefinition, "Pivot2");
            p2.ColumnFields.Add(p2.Fields[1]);
            p2.DataFields.Add(p2.Fields[4]);

            Assert.AreEqual(5, p1.CacheDefinition._cacheReference.Fields.Count);
            Assert.AreEqual(p1.CacheDefinition._cacheReference, p2.CacheDefinition._cacheReference);
        }
        [TestMethod]
        public void ValidateDifferentChangeToSameCache()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotChangeCache");
            var p1 = ws.PivotTables.Add(ws.Cells["A1"], _pck.Workbook.Worksheets[0].Tables[0].Range, "Pivot3");
            p1.RowFields.Add(p1.Fields[0]);
            p1.DataFields.Add(p1.Fields[3]);
            var p2 = ws.PivotTables.Add(ws.Cells["K1"], _pck.Workbook.Worksheets[1].Tables[0].Range, "Pivot4");
            p2.ColumnFields.Add(p2.Fields[1]);
            p2.DataFields.Add(p2.Fields[4]);

            Assert.AreEqual(5, p1.CacheDefinition._cacheReference.Fields.Count);
//            Assert.AreEqual(2, _pck.Workbook._pivotTableCaches.Count);

            p2.CacheDefinition.SourceRange = _pck.Workbook.Worksheets[0].Tables[0].Range;
        }
        [TestMethod]
        public void ValidateSameCacheThenNewCache()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotSameThenNewCache");
            var p1 = ws.PivotTables.Add(ws.Cells["A1"], _pck.Workbook.Worksheets[0].Tables[0].Range, "Pivot5");
            p1.RowFields.Add(p1.Fields[0]);
            p1.DataFields.Add(p1.Fields[3]);
            var p2 = ws.PivotTables.Add(ws.Cells["K1"], _pck.Workbook.Worksheets[0].Tables[0].Range, "Pivot6");
            p2.ColumnFields.Add(p2.Fields[1]);
            p2.DataFields.Add(p2.Fields[4]);

            Assert.AreEqual(5, p1.CacheDefinition._cacheReference.Fields.Count);

            p2.CacheDefinition.SourceRange = _pck.Workbook.Worksheets[1].Tables[0].Range;
        }

        [TestMethod]
        public void ValidateSameCacheDateGrouping()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotSameCacheDateGroup");
            var p1 = ws.PivotTables.Add(ws.Cells["A1"], _pck.Workbook.Worksheets[0].Tables[0].Range, "Pivot1");
            p1.RowFields.Add(p1.Fields[0]);
            p1.DataFields.Add(p1.Fields[3]);
            var p2 = ws.PivotTables.Add(ws.Cells["K1"], p1.CacheDefinition, "Pivot2");
            p2.DataFields.Add(p2.Fields[3]);
            p2.RowFields.Add(p2.Fields[4]);
            p2.Fields[4].AddDateGrouping(eDateGroupBy.Years | eDateGroupBy.Months | eDateGroupBy.Days);

            Assert.AreEqual(7, p1.CacheDefinition._cacheReference.Fields.Count);
            Assert.AreEqual(p1.CacheDefinition._cacheReference, p2.CacheDefinition._cacheReference);
        }
    }
}
