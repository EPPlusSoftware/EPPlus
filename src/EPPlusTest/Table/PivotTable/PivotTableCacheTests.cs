using FakeItEasy.Configuration;
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
        [TestMethod]
        public void ValidateTimeSpanHandligInCache()
        {
            ExcelWorksheet wsData = _pck.Workbook.Worksheets.Add("Data");
            wsData.Column(2).Style.Numberformat.Format = "m/d/yyyy";
            wsData.Column(2).Width = 12;
            wsData.Column(3).Style.Numberformat.Format = "HH:MM:SS";

            wsData.Cells["A1"].Value = "Text";
            wsData.Cells["B1"].Value = "Date";
            wsData.Cells["C1"].Value = "Time";

            wsData.Cells["A2"].Value = "Row1";
            wsData.Cells["B2"].Value = DateTime.Today;
            wsData.Cells["C2"].Value = new TimeSpan(500);

            wsData.Cells["A3"].Value = "Row2";
            wsData.Cells["B3"].Value = DateTime.Today;
            wsData.Cells["C3"].Value = new TimeSpan(7000000);

            ExcelWorksheet wsPivot = _pck.Workbook.Worksheets.Add("PivotDateAndTimeSpan");
            var dataRange = wsData.Cells[wsData.Dimension.Address.ToString()];
            dataRange.AutoFitColumns();
            var pivotTable = wsPivot.PivotTables.Add(wsPivot.Cells["A3"], dataRange, "Pivotname");
            pivotTable.MultipleFieldFilters = true;
            pivotTable.RowGrandTotals = true;
            pivotTable.ColumnGrandTotals = true;
            pivotTable.Compact = true;
            pivotTable.CompactData = true;
            pivotTable.GridDropZones = false;
            pivotTable.Outline = false;
            pivotTable.OutlineData = false;
            pivotTable.ShowError = true;
            pivotTable.ErrorCaption = "[error]";
            pivotTable.ShowHeaders = true;
            pivotTable.UseAutoFormatting = true;
            pivotTable.ApplyWidthHeightFormats = true;
            pivotTable.ShowDrill = true;
            pivotTable.FirstDataCol = 3;
            pivotTable.RowHeaderCaption = "Date";

            var dateField = pivotTable.Fields["Date"];
            pivotTable.RowFields.Add(dateField);



            var timeField = pivotTable.Fields["Time"];
            pivotTable.RowFields.Add(timeField);
            timeField.Cache.Refresh();
            Assert.AreEqual(2, timeField.Cache.SharedItems.Count);
            Assert.AreEqual(new DateTime(0), timeField.Cache.SharedItems[0]);
            Assert.AreEqual(new DateTime(TimeSpan.TicksPerSecond), timeField.Cache.SharedItems[1]);

            var countField = pivotTable.Fields["Text"];
            pivotTable.ColumnFields.Add(countField);
        }
        [TestMethod]
        public void ValidatePivotTableCacheAfterDeletedWorksheet()
        {
            using (var p1 = new ExcelPackage())
            {
                ExcelWorksheet wsData = p1.Workbook.Worksheets.Add("DataDeleted");
                ExcelWorksheet wsPivot = p1.Workbook.Worksheets.Add("PivotDeleted");
                LoadTestdata(wsData);
                var dataRange = wsData.Cells[wsData.Dimension.Address.ToString()];
                var pivotTable = wsPivot.PivotTables.Add(wsPivot.Cells["A3"], dataRange, "PivotDeleted");
                p1.Save();
                using(var p2=new ExcelPackage(p1.Stream))
                {
                    p2.Workbook.Worksheets.Delete("DataDeleted");
                    wsData = p2.Workbook.Worksheets.Add("DataDeleted");
                    LoadTestdata(wsData);

                    SaveWorkbook("pivotDeletedWorksheet.xlsx", p2);
                }
            }

        }
    }
}
