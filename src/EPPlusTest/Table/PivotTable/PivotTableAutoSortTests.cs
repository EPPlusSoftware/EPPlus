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
    public class PivotTableAutoSortTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("PivotTableAutoSort.xlsx", true);
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
        public void SetAutoSortAcending()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotSameAutoSortAcending");
            var p1 = ws.PivotTables.Add(ws.Cells["A1"], _pck.Workbook.Worksheets[0].Tables[0].Range, "Pivot1");
            var rf=p1.RowFields.Add(p1.Fields[0]);
            var df=p1.DataFields.Add(p1.Fields[3]);
            rf.SetAutoSort(df);
        }
        [TestMethod]
        public void SetAutoSortDesending()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotSameAutoSortDescending");
            var p1 = ws.PivotTables.Add(ws.Cells["A1"], _pck.Workbook.Worksheets[0].Tables[0].Range, "Pivot2");
            var rf = p1.RowFields.Add(p1.Fields[0]);
            var df = p1.DataFields.Add(p1.Fields[3]);
            rf.SetAutoSort(df, eSortType.Descending);
        }
        [TestMethod]
        public void SetAutoSortDataAndColumnField1()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotSameAutoSortDescendingCF1");
            var p1 = ws.PivotTables.Add(ws.Cells["A1"], _pck.Workbook.Worksheets[0].Tables[0].Range, "Pivot2");            
            var rf = p1.RowFields.Add(p1.Fields[0]);            
            var cf = p1.ColumnFields.Add(p1.Fields[1]);
            var df = p1.DataFields.Add(p1.Fields[3]);
            rf.SetAutoSort(df, eSortType.Descending);
            var reference = rf.AutoSort.Conditions.Fields.Add(cf);
            cf.Items.Refresh();
            reference.Items.AddByValue("Hardware");
        }
        [TestMethod]
        public void SetAutoSortDataAndColumnField2()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotSameAutoSortDescendingCF2");
            var p1 = ws.PivotTables.Add(ws.Cells["A1"], _pck.Workbook.Worksheets[0].Tables[0].Range, "Pivot2");
            var rf = p1.RowFields.Add(p1.Fields[0]);
            var cf = p1.ColumnFields.Add(p1.Fields[1]);
            var df = p1.DataFields.Add(p1.Fields[3]);
            rf.SetAutoSort(df, eSortType.Descending);
            var reference = rf.AutoSort.Conditions.Fields.Add(cf);
            cf.Items.Refresh();
            reference.Items.Add(1);
        }
        [TestMethod]
        public void SetAutoSortDataAndRowField1()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotSameAutoSortDescendingRF1");
            var p1 = ws.PivotTables.Add(ws.Cells["A1"], _pck.Workbook.Worksheets[0].Tables[0].Range, "Pivot2");
            var rf = p1.RowFields.Add(p1.Fields[0]);
            var cf = p1.ColumnFields.Add(p1.Fields[1]);
            var df = p1.DataFields.Add(p1.Fields[3]);
            cf.SetAutoSort(df, eSortType.Descending);
            var reference = cf.AutoSort.Conditions.Fields.Add(rf);
            rf.Items.Refresh();
            reference.Items.Add(0);
        }
        [TestMethod]
        public void SetAutoSortDataAndRowField3()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotSameAutoSortDescendingRF3");
            var p1 = ws.PivotTables.Add(ws.Cells["A1"], _pck.Workbook.Worksheets[0].Tables[0].Range, "Pivot2");
            var rf = p1.RowFields.Add(p1.Fields[0]);
            var cf = p1.ColumnFields.Add(p1.Fields[1]);
            var df = p1.DataFields.Add(p1.Fields[3]);
            cf.SetAutoSort(df, eSortType.Descending);
            var reference = cf.AutoSort.Conditions.Fields.Add(rf);
            rf.Items.Refresh();
            reference.Items.Add(2);
        }
        [TestMethod]
        public void ReadAutoSort()
        {
            using(var p1=new ExcelPackage())
            {
                var ws = p1.Workbook.Worksheets.Add("PivotSameAutoClear");
                var r = LoadItemData(ws);
                ws.Tables.Add(r, "Table1");

                var pivot1 = ws.PivotTables.Add(ws.Cells["A1"], p1.Workbook.Worksheets[0].Tables[0].Range, "Pivot2");
                var rf = pivot1.RowFields.Add(pivot1.Fields[0]);
                var cf = pivot1.ColumnFields.Add(pivot1.Fields[1]);
                var df = pivot1.DataFields.Add(pivot1.Fields[3]);
                cf.SetAutoSort(df, eSortType.Descending);
                var reference = cf.AutoSort.Conditions.Fields.Add(rf);
                rf.Items.Refresh();
                reference.Items.Add(2);

                Assert.IsNotNull(cf.AutoSort);

                p1.Save();

                using(var p2=new ExcelPackage(p1.Stream))
                {
                    var ws1 = p1.Workbook.Worksheets[0];
                    var pivot2 = ws.PivotTables[0];

                    Assert.AreEqual(1, pivot2.ColumnFields.Count);
                    Assert.AreEqual(1, pivot2.RowFields.Count);
                    Assert.AreEqual(1, pivot2.DataFields.Count);
                    Assert.IsNotNull(pivot2.ColumnFields[0].AutoSort);
                    Assert.AreEqual(1, pivot2.ColumnFields[0].AutoSort.Conditions.DataFields.Count);
                    Assert.AreEqual(1, pivot2.ColumnFields[0].AutoSort.Conditions.Fields.Count);
                }

            }
        }
        [TestMethod]
        public void RemoveAutoSort()
        {
            using (var p1 = new ExcelPackage())
            {
                var ws = p1.Workbook.Worksheets.Add("PivotSameAutoClear");
                var r = LoadItemData(ws);
                ws.Tables.Add(r, "Table1");

                var pivot1 = ws.PivotTables.Add(ws.Cells["A1"], p1.Workbook.Worksheets[0].Tables[0].Range, "Pivot2");
                var rf = pivot1.RowFields.Add(pivot1.Fields[0]);
                var cf = pivot1.ColumnFields.Add(pivot1.Fields[1]);
                var df = pivot1.DataFields.Add(pivot1.Fields[3]);
                cf.SetAutoSort(df, eSortType.Descending);
                var reference = cf.AutoSort.Conditions.Fields.Add(rf);
                rf.Items.Refresh();
                reference.Items.Add(2);

                Assert.IsNotNull(cf.AutoSort);

                p1.Save();

                using (var p2 = new ExcelPackage(p1.Stream))
                {
                    var ws1 = p1.Workbook.Worksheets[0];
                    var pivot2 = ws.PivotTables[0];

                    Assert.AreEqual(1, pivot2.ColumnFields.Count);
                    Assert.AreEqual(1, pivot2.RowFields.Count);
                    Assert.AreEqual(1, pivot2.DataFields.Count);
                    Assert.IsNotNull(pivot2.ColumnFields[0].AutoSort);

                    pivot2.ColumnFields[0].RemoveAutoSort();
                    Assert.IsNull(pivot2.ColumnFields[0].AutoSort);
                }

            }
        }
    }
}
