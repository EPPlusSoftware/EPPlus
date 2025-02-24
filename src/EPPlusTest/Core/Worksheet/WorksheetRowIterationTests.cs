using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace EPPlusTest.Core.Worksheet
{
    [TestClass]
    public class WorksheetRowIterationTests : TestBase
    {
        [TestMethod]
        public void RowsIterateEmptyShouldSkip()
        {
            using (var p = OpenPackage("DeleteRows.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("rowSheet");

                int iteratedRows = 0;

                var rows = ws.Rows;
                foreach (var row in rows)
                {
                    if (row.StartRow % 2 > 0)
                    {
                        ws.DeleteRow(row.StartRow);
                    }
                    iteratedRows++;
                }

                Assert.AreEqual(0, iteratedRows);
            }
        }

        [TestMethod]
        public void EnsureEntireRowWorks()
        {
            using (var p = OpenPackage("DeleteRows.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("rowSheet");

                var range = ws.Cells["A1:D10"];

                range.Value = "";

                range.EntireRow.Style.Fill.SetBackground(Color.DarkRed);

                var royalBlue = ws.Cells["A1"].Style.Fill.BackgroundColor.Rgb;

                var subRange = ws.Cells["A3:D6"];

                var subRows = subRange.EntireRow;
                subRows.Style.Fill.SetBackground(Color.Yellow);

                var yellowRgb = subRows.Style.Fill.BackgroundColor.Rgb;

                foreach (var row in subRows)
                {
                    Assert.AreEqual(yellowRgb, row.Style.Fill.BackgroundColor.Rgb);
                }

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void DeleteAllWithPredicate()
        {
            using (var p = OpenPackage("DeleteRowsWithPredicate.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("rowSheet");

                var range = ws.Cells["A1:D10"];

                range.Value = "";

                range.EntireRow.Style.Fill.SetBackground(Color.DarkRed);

                var royalBlue = ws.Cells["A1"].Style.Fill.BackgroundColor.Rgb;

                var subRange = ws.Cells["A3:D6"];

                var subRows = subRange.EntireRow;
                subRows.Style.Fill.SetBackground(Color.Yellow);

                var yellowRgb = subRows.Style.Fill.BackgroundColor.Rgb;

                ws.Rows.DeleteAll(row => row.Style.Fill.BackgroundColor.Rgb == yellowRgb);

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void IterationTest()
        {
            using (var p = OpenPackage("IterateRows.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("rowSheet");

                ws.Cells[4, 2].Value = 0;

                ws.Row(5).Hidden = true;
                ws.Row(6).Hidden = true;

                ws.Cells[7, 2].Value = 5;

                ws.Row(9).Hidden = true;
                ws.Cells[10, 2].Value = 10;

                ws.Rows.DeleteAll(r => r.Hidden == true);

                //Assert.AreEqual(0, ws.Cells[4, 2].Value);
                //Assert.AreEqual(5, ws.Cells[7, 2].Value);
                //Assert.AreEqual(10, ws.Cells[9, 2].Value);


                //var hiddenRows = ws.Rows.Where(r => r.Hidden == true);

                //foreach (var row in hiddenRows)
                //{
                //    ws.DeleteRow(row.StartRow);
                //}

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void IterationTest2()
        {
            using (var p = OpenPackage("IterateRows2.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("rowSheet");

                ws.Row(2).Height = ws.Row(2).Height * 1.1;
                ws.Row(5).Hidden = true;
                ws.Row(6).Hidden = true;

                ws.Cells[7, 2].Value = 5;

                ws.Row(9).Hidden = true;
                ws.Cells[10, 2].Value = 10;


                var hiddenRows = ws.Rows.Where(r => r.Hidden == true);

                List<int> startingRows = new();

                foreach (var row in hiddenRows)
                {
                    startingRows.Add(row.StartRow);
                }

                Assert.AreEqual(3, startingRows.Count());
                Assert.AreEqual(5, startingRows[0]);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void UnsetRowsIteration()
        {
            using (var p = OpenPackage("UnsetRows.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("rowSheet");


                var range = ws.Cells["N1:R5"];


                int nrRowsBefore = 0;
                foreach (var row in ws.Rows[range.Start.Row, range.End.Row])
                {
                    row.Style.Fill.SetBackground(Color.Red);
                    nrRowsBefore++;
                }

                range.Value = "1";

                var rows = ws.Rows[range.Start.Row, range.End.Row];

                int nrRowsAfter = 0;
                foreach (var row in rows)
                {
                    row.Style.Fill.SetBackground(Color.Red);
                    nrRowsAfter++;
                }

                //ws.Row(2).Height = ws.Row(2).Height * 1.1;
                //ws.Row(5).Hidden = true;
                //ws.Row(6).Hidden = true;

                //ws.Cells[7, 2].Value = 5;

                //ws.Row(9).Hidden = true;
                //ws.Cells[10, 2].Value = 10;


                //var hiddenRows = ws.Rows.Where(r => r.Hidden == true);

                //List<int> startingRows = new();

                //foreach (var row in hiddenRows)
                //{
                //    startingRows.Add(row.StartRow);
                //}

                //Assert.AreEqual(3, startingRows.Count());
                //Assert.AreEqual(5, startingRows[0]);

                SaveAndCleanup(p);
            }
        }
    }
}
