using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
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

                var range = ws.Cells["A2:B11"];
                var rows = ws.Rows[range.Start.Row, range.End.Row];

                List<int> otherRows = new();
                foreach (var row in rows)
                {
                    otherRows.Add(row.StartRow);
                }

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
        public void NumRowsSetRowColStyles()
        {
            using (var p = OpenPackage("NumRows_SetRowColStyles.xlsx", true))
            {
                //List<int> inmts = new List<int> { 1, 2, 3, 4 };

                //int numThrough = 0;
                //foreach(var num in inmts)
                //{
                //    numThrough++;
                //}

                //foreach (var num in inmts)
                //{
                //    numThrough++;
                //}

                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("rowSheet");

                var values = ws._values;

                ws.Cells["A1"].EntireColumn.Style.Fill.SetBackground(Color.Blue);

                var range = ws.Cells["N1:R5"];

                range.Value = "1";

                var rows = ws.Rows[range.Start.Row, range.End.Row];

                var pageRows = ws._values._columnIndex[0]._pages[0].Rows;

                var pageRowsAfter = ws._values._columnIndex[1]._pages[0].Rows;

                int nrRowsAfter = 0;
                foreach (var row in rows)
                {
                    nrRowsAfter++;
                }

                foreach (var row in rows)
                {
                    row.Style.Fill.SetBackground(Color.Red);
                    nrRowsAfter++;
                }

                SaveAndCleanup(p);

                Assert.AreEqual(10, nrRowsAfter);
            }
        }

        [TestMethod]
        public void NumRowsSetRowsAfterCols()
        {
            using (var p = OpenPackage("NumRows_SetRowsAfterCols.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("iterateSetWs");

                var range = ws.Cells["N1:R5"];
                range.Value = "1";

                var rows = ws.Rows[range.Start.Row, range.End.Row];
                var cols = ws.Columns[range.Start.Column, range.End.Column];

                int colNr = 0;
                int nrRowsAfter = 0;
                foreach (var col in cols)
                {
                    colNr++;
                }

                foreach (var row in rows)
                {
                    row.Style.Fill.SetBackground(Color.Red);
                    nrRowsAfter++;
                }

                Assert.AreEqual(5, colNr);
                Assert.AreEqual(5, nrRowsAfter);

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


                //int nrRowsBefore = 0;
                //foreach (var row in ws.Rows[range.Start.Row, range.End.Row])
                //{
                //    row.Style.Fill.SetBackground(Color.Red);
                //    nrRowsBefore++;
                //}

                //int nrColsBefore = 0;

                //foreach (var col in ws.Columns[range.Start.Column, range.End.Column])
                //{
                //    ws.Columns[col.StartColumn].Style.Fill.SetBackground(Color.Red);
                //    nrColsBefore++;
                //}

                range.Value = "1";

                var rows = ws.Rows[range.Start.Row, range.End.Row];
                //var cols = ws.Columns[range.Start.Column, range.End.Column];

                int nrRowsAfter = 0;
                foreach (var row in rows)
                {
                    row.Style.Fill.SetBackground(Color.Red);
                    nrRowsAfter++;
                }

                var entireRows = range.EntireRow;
                int nrEntireRowsAfter = 0;

                foreach (var row in entireRows)
                {
                    if (row.StartRow % 2 > 0)
                    {
                        row.Style.Fill.SetBackground(Color.Red);
                    }
                    else
                    {
                        row.Style.Fill.SetBackground(Color.Blue);
                    }
                    nrEntireRowsAfter++;
                }

                int nrColsAfter = 0;
                foreach (var col in ws.Columns[range.Start.Column, range.End.Column])
                {
                    ws.Columns[col.StartColumn].Style.Fill.SetBackground(Color.Red);
                    nrColsAfter++;
                }

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void SomeColIssues()
        {
            using (var p = OpenTemplatePackage("ColRowIssueWCells.xlsx"))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets[0];

                var range = ws.Cells["C1:G5"];
                var rows = ws.Rows[range.Start.Row, range.End.Row];

                int rowNr = 0;
                foreach (var row in rows)
                {
                    rowNr++;
                }

                Assert.AreEqual(5, rowNr);
            }
        }

        [TestMethod]
        public void SomeColIssues2()
        {
            using (var p = OpenTemplatePackage("ColRowIssue.xlsx"))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets[0];

                var range = ws.Cells["C1:F6"];
                var rows = ws.Rows[range.Start.Row, range.End.Row];

                int rowNr = 0;
                foreach (var row in rows)
                {
                    rowNr++;
                }

                Assert.AreEqual(6, rowNr);
            }
        }
        [TestMethod]
        public void RowOnlyLab()
        {
            using (var p = OpenTemplatePackage("RowOnlyLab.xlsx"))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets[0];

                int rowNr = 0;
                var rows = ws.Rows;

                var startRow = ws.Rows.StartRow;
                var endRow = ws.Rows.EndRow;

                //var dimensionRows = ws.Rows[ws.Dimension.Start.Row,ws.Dimension.End.Row];

                //int numDimRows = 0;
                //foreach(var dimRow in dimensionRows)
                //{
                //    numDimRows++;
                //}

                foreach (var row in rows)
                {
                    rowNr++;
                }

                Assert.AreEqual(7, rowNr);
            }
        }

        [TestMethod]
        public void RowsAndColumnsWillnotIterateIfUnset()
        {
            using (var p = OpenPackage("NoIterate.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("noIterateWs");

                var range = ws.Cells["N1:R5"];

                var rows = ws.Rows[range.Start.Row, range.End.Row];
                var cols = ws.Columns[range.Start.Column, range.End.Column];

                int colNr = 0;
                int rowNr = 0;
                foreach(var col in cols)
                {
                    colNr++;
                }

                foreach(var row in rows)
                {
                    rowNr++;
                }

                //This seems like intended behaviour as no data is actually on rows columns or cells
                //It is still somewhat confusing to an end user.
                Assert.AreEqual(0, colNr);
                Assert.AreEqual(0, rowNr);

                var entireColumns = range.EntireColumn;
                var entireRows = range.EntireRow;

                foreach(var col in entireColumns)
                {
                    colNr++;
                }

                foreach(var row in entireRows)
                {
                    rowNr++;
                }

                Assert.AreEqual(0, colNr);
                Assert.AreEqual(0, rowNr);

                range.EntireColumn.Hidden = true;
                range.EntireRow.Hidden = true;

                var entireColumns2 = range.EntireColumn;
                foreach (var col in entireColumns2)
                {
                    colNr++;
                }

                var entireRows2 = range.EntireRow;
                foreach (var row in entireRows2)
                {
                    rowNr++;
                }

                SaveAndCleanup(p);

                Assert.AreEqual(5, colNr);
                Assert.AreEqual(5, rowNr);
            }
        }

        [TestMethod]
        public void RowsAndColumnsWillIterateIfSet()
        {
            using (var p = OpenPackage("iterateSet.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("iterateSetWs");

                var range = ws.Cells["N1:R5"];
                range.Value = "1";

                var rows = ws.Rows[range.Start.Row, range.End.Row];
                var cols = ws.Columns[range.Start.Column, range.End.Column];

                int colNr = 0;
                int rowNr = 0;
                foreach (var col in cols)
                {
                    colNr++;
                }

                foreach (var row in rows)
                {
                    rowNr++;
                }

                //This seems like intended behaviour as no data is actually on rows columns or cells
                //It is still somewhat confusing to an end user.
                Assert.AreEqual(5, colNr);
                Assert.AreEqual(5, rowNr);

                var entireColumns = range.EntireColumn;
                var entireRows = range.EntireRow;

                foreach (var col in entireColumns)
                {
                    colNr++;
                }

                foreach (var row in entireRows)
                {
                    rowNr++;
                }

                Assert.AreEqual(10, colNr);
                Assert.AreEqual(10, rowNr);

                range.EntireColumn.Hidden = true;
                range.EntireRow.Hidden = true;

                var entireColumns2 = range.EntireColumn;
                foreach (var col in entireColumns2)
                {
                    colNr++;
                }

                var entireRows2 = range.EntireRow;
                foreach (var row in entireRows2)
                {
                    rowNr++;
                }

                Assert.AreEqual(15, colNr);
                Assert.AreEqual(15, rowNr);
            }
        }

        [TestMethod]
        public void DoesNotCareAboutSetCellsOnRows()
        {
            using (var p = OpenPackage("IteratingAndAddingToRows.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("iterateSetWs");

                var range = ws.Cells["N1:R5"];
                range.Value = "1";

                var rows = ws.Rows[range.Start.Row, range.End.Row];

                foreach (var row in rows)
                {
                    row.Style.Fill.SetBackground(Color.Blue);
                    ws.InsertRow(row.StartRow + 1, 1);
                }

                SaveAndCleanup(p);
            }
        }
    }
}
