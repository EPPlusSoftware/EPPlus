
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Table.PivotTable;
using System;

namespace EPPlusTest.Table.PivotTable
{
    [TestClass]
    public class PivotTableDateFilterTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("PivotTableDateFilters.xlsx", true);
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
        public void AddDateEqualFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("DateEqual");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateEqual, new DateTime(2010,3,31));
        }
        [TestMethod]
        public void AddDateNotEqualFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("DateNotEqual");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateNotEqual, new DateTime(2010, 3, 31));
        }
        [TestMethod]
        public void AddDateOlderFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("DateBefore");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateOlderThan, new DateTime(2010, 3, 31));
        }
        [TestMethod]
        public void AddDateOlderOrEqualFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("DateBeforeOrEqual");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateOlderThanOrEqual, new DateTime(2010, 3, 31));
        }
        [TestMethod]
        public void AddDateNewerFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("DateNewer");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateNewerThan, new DateTime(2010, 3, 31));
        }
        [TestMethod]
        public void AddDateNewerOrEqualFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("DateNewerOrEqual");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateNewerThanOrEqual, new DateTime(2010, 3, 31));
        }
        [TestMethod]
        public void AddDateBetweenFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("DateBetween");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateBetween, new DateTime(2010, 3, 31), new DateTime(2010, 6, 30));
        }
        [TestMethod]
        public void AddDateNotBetweenFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("DateNotBetween");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateNotBetween, new DateTime(2010, 3, 31), new DateTime(2010, 6, 30));
        }
        [TestMethod]
        public void AddDateLastMonthFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("LastMonth");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.LastMonth);
        }
        [TestMethod]
        public void AddDateLastQuarterFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("LastQuarter");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.LastQuarter);
        }
        [TestMethod]
        public void AddDateLastWeekFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("LastWeek");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.LastWeek);
        }
        [TestMethod]
        public void AddDateLastYearFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("LastYear");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.LastYear);
        }
        [TestMethod]
        public void AddDateM1Filter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("M1");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M1);
        }
        [TestMethod]
        public void AddDateM2Filter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("M2");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M2);
        }
        [TestMethod]
        public void AddDateM3Filter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("M3");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M3);
        }
        [TestMethod]
        public void AddDate42Filter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("M4");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M4);
        }
        [TestMethod]
        public void AddDateM5Filter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("M5");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M5);
        }
        [TestMethod]
        public void AddDateM6Filter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("M6");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M6);
        }
        [TestMethod]
        public void AddDateM7Filter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("M7");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M7);
        }
        [TestMethod]
        public void AddDateM8Filter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("M8");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M8);
        }
        [TestMethod]
        public void AddDateM9Filter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("M9");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M9);
        }
        [TestMethod]
        public void AddDateM10Filter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("M10");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M10);
        }
        [TestMethod]
        public void AddDateM11Filter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("M11");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M11);
        }
        [TestMethod]
        public void AddDateM12Filter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("M12");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.M12);
        }
        [TestMethod]
        public void AddDateQ1Filter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("Q1");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.Q1);
        }
        [TestMethod]
        public void AddDateQ2Filter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("Q2");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.Q2);
        }
        [TestMethod]
        public void AddDateQ3Filter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("Q3");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.Q3);
        }
        [TestMethod]
        public void AddDateQ4Filter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("Q4");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.Q4);
        }
        [TestMethod]
        public void AddDateYesterdayFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("Yesterday");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.Yesterday);
        }
        [TestMethod]
        public void AddDateTodayFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("Today");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.Today);
        }
        [TestMethod]
        public void AddDateTomorrowFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("Tomorrow");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.Tomorrow);
        }
        [TestMethod]
        public void AddDateYTDFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("YTD");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.YearToDate);
        }
        [TestMethod]
        public void AddDateThisMonthFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("ThisMonth");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.ThisMonth);
        }
        [TestMethod]
        public void AddDateThisQuarterFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("ThisQuarter");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.ThisQuarter);
        }
        [TestMethod]
        public void AddDateThisWeekFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("ThisWeek");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.ThisWeek);
        }
        [TestMethod]
        public void AddDateThisYearFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("ThisYear");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.ThisYear);
        }
        [TestMethod]
        public void AddDateNextMonthFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("NextMonth");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.NextMonth);
        }
        [TestMethod]
        public void AddDateNextQuarterFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("NextQuarter");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.NextQuarter);
        }
        [TestMethod]
        public void AddDateNextWeekFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("NextWeek");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.NextWeek);
        }
        [TestMethod]
        public void AddDateNextYearFilter()
        {
            var wsData = _pck.Workbook.Worksheets["Data1"];
            var ws = _pck.Workbook.Worksheets.Add("NextYear");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataFields.Add(pt.Fields[3]);

            pt.Fields[4].Filters.AddDatePeriodFilter(ePivotTableDatePeriodFilterType.NextYear);
        }
    }
}
