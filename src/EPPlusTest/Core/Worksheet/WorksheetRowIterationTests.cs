using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Utilities;
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
        public void IterateRowsWithPropeties()
        {
            using (var p = OpenPackage("Rows_Iterate_WithProperty.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("rowSheet");

                int iteratedRows = 0;

                ws.Row(1).Hidden = true;
                ws.Row(2).Hidden = true;
                ws.Row(3).Hidden = true;

                var rows = ws.Rows;
                foreach (var row in rows)
                {
                    Assert.IsTrue(row.Hidden);
                    iteratedRows++;
                }

                Assert.AreEqual(3, iteratedRows);
                SaveAndCleanup(p);
            }
        }


        [TestMethod]
        public void DoNotIterateEmpty()
        {
            using (var p = OpenPackage("IterateRows_Empty.xlsx", true))
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
        public void DeleteAllRowslWithPredicate()
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
        public void DeleteAllHiddenRows()
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

                Assert.AreEqual(0, ws.Cells[4, 2].Value);
                Assert.AreEqual(5, ws.Cells[5, 2].Value);
                Assert.AreEqual(10, ws.Cells[7, 2].Value);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void SetStylingOnRows()
        {
            using (var p = OpenPackage("IterateRows_SetRowColStyles.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("rowSheet");

                var values = ws._values;

                ws.Cells["A1"].EntireColumn.Style.Fill.SetBackground(Color.Blue);

                var range = ws.Cells["N1:R5"];

                range.Value = "1";

                var rows = ws.Rows[range.Start.Row, range.End.Row];

                int nrRowsAfter = 0;
                foreach (var row in rows)
                {
                    row.Style.Fill.SetBackground(Color.IndianRed);
                    nrRowsAfter++;
                }

                Assert.AreEqual(5, nrRowsAfter);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void RowsIterateEntireRowSubRange()
        {
            using (var p = OpenPackage("IterateRows_EntireRowSubrange.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("rowSheet");

                var range = ws.Cells["A1:D10"];

                range.Value = "";

                range.EntireRow.Style.Fill.SetBackground(Color.DarkRed);

                var darkRed = ws.Cells["A1"].Style.Fill.BackgroundColor.Rgb;

                var subRange = ws.Cells["A3:D6"];

                var subRows = subRange.EntireRow;
                subRows.Style.Fill.SetBackground(Color.Yellow);

                var yellowRgb = subRows.Style.Fill.BackgroundColor.Rgb;

                var blueRgb = "";

                List<int> iteratedRowIndicies = new();

                foreach (var row in subRows)
                {
                    iteratedRowIndicies.Add(row.StartRow);
                    Assert.AreEqual(yellowRgb, row.Style.Fill.BackgroundColor.Rgb);
                    if (row.StartRow == 3 || row.StartRow == 5)
                    {
                        row.Style.Fill.SetBackground(Color.Blue);
                        blueRgb = row.Style.Fill.BackgroundColor.Rgb;
                    }
                }

                Assert.AreEqual(blueRgb, ws.Cells["A3"].EntireRow.Style.Fill.BackgroundColor.Rgb);
                Assert.AreEqual(blueRgb, ws.Cells["A5"].EntireRow.Style.Fill.BackgroundColor.Rgb);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void IterateRowsWithValues()
        {
            using (var p = OpenPackage("IterateRows_WithValues.xlsx", true))
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

                List<int> iteratedValueRows = new();
                foreach (var row in rows)
                {
                    iteratedValueRows.Add(row.StartRow);
                }
                Assert.AreEqual(6, iteratedValueRows.Count);

                Assert.AreEqual(2, iteratedValueRows[0]);
                Assert.AreEqual(5, iteratedValueRows[1]);
                Assert.AreEqual(6, iteratedValueRows[2]);
                Assert.AreEqual(7, iteratedValueRows[3]);
                Assert.AreEqual(9, iteratedValueRows[4]);
                Assert.AreEqual(10, iteratedValueRows[5]);


                var hiddenRows = ws.Rows.Where(r => r.Hidden == true);

                List<int> foundRows = new();

                foreach (var row in hiddenRows)
                {
                    foundRows.Add(row.StartRow);
                }

                Assert.AreEqual(3, foundRows.Count());

                Assert.AreEqual(5, foundRows[0]);
                Assert.AreEqual(6, foundRows[1]);
                Assert.AreEqual(9, foundRows[2]);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void NumRowsSetRowColStyles()
        {
            using (var p = OpenPackage("NumRows_SetRowColStyles.xlsx", true))
            {
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

                range.Value = "1";

                var rows = ws.Rows[range.Start.Row, range.End.Row];

                int nrRowsAfter = 0;
                foreach (var row in rows)
                {
                    row.Style.Fill.SetBackground(Color.Yellow);
                    nrRowsAfter++;
                }

                Assert.AreEqual(5, nrRowsAfter);

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

                Assert.AreEqual(5, nrEntireRowsAfter);

                int nrColsAfter = 0;
                foreach (var col in ws.Columns[range.Start.Column, range.End.Column])
                {
                    ws.Columns[col.StartColumn].Style.Fill.SetBackground(Color.Red);
                    nrColsAfter++;
                }

                Assert.AreEqual(5, nrColsAfter);

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
                Assert.AreEqual(0, colNr);
                Assert.AreEqual(0, rowNr);

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

                Assert.AreEqual(0, colNr);
                Assert.AreEqual(0, rowNr);

                range.EntireColumn.Hidden = true;
                range.EntireRow.Hidden = true;

                var entireColumns2 = range.EntireColumn;
                foreach (var col in entireColumns2)
                {
                    colNr++;
                }

                List<int> iteratedRows = new();

                var entireRows2 = range.EntireRow;
                foreach (var row in entireRows2)
                {
                    iteratedRows.Add(row.StartRow);
                    rowNr++;
                }

                SaveAndCleanup(p);

                Assert.AreEqual(5, colNr);
                Assert.AreEqual(5, rowNr);

                Assert.AreEqual(5, iteratedRows.Count());
                Assert.AreEqual(1, iteratedRows[0]);
                Assert.AreEqual(2, iteratedRows[1]);
                Assert.AreEqual(3, iteratedRows[2]);
                Assert.AreEqual(4, iteratedRows[3]);
                Assert.AreEqual(5, iteratedRows[4]);
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
        public void FindsRowsIfCellValueSet()
        {
            using (var p = OpenPackage("FindsRowsIfCellValueSet.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("iterateSetWs");

                var range = ws.Cells["N1:R5"];

                ws.Cells["N2"].Value = 5;

                ws.Cells["R4"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Dashed;

                ws.Row(5).Hidden = true;

                List<int> iteratedRows = new();

                foreach (var row in range.EntireRow)
                {
                    iteratedRows.Add(row.StartRow);
                    row.Style.Fill.SetBackground(Color.CadetBlue);
                }

                Assert.AreEqual(3, iteratedRows.Count());
                Assert.AreEqual(2, iteratedRows[0]);
                Assert.AreEqual(4, iteratedRows[1]);
                Assert.AreEqual(5, iteratedRows[2]);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void IteratingAndAddingToRowsRange()
        {
            using (var p = OpenPackage("IteratingAndAddingToRowsRange.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("iterateSetWs");

                var range = ws.Cells["N1:R5"];


                ws.Cells["N2"].Value = 5;

                ws.Cells["R4"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Dashed;

                ws.Row(5).Hidden = true;

                List<int> iteratedRows = new();

                foreach (var row in range.EntireRow)
                {
                    iteratedRows.Add(row.StartRow);
                    row.Style.Fill.SetBackground(Color.CadetBlue);
                }

                Assert.AreEqual(3, iteratedRows.Count());
                Assert.AreEqual(2, iteratedRows[0]);
                Assert.AreEqual(4, iteratedRows[1]);
                Assert.AreEqual(5, iteratedRows[2]);

                var rows = ws.Rows[range.Start.Row, range.End.Row];

                int iteratedNr = 0;

                foreach (var row in rows)
                {
                    ws.InsertRow(row.StartRow + 1, 1);
                    ws.Row(row.StartRow + 1).Height = ws.Row(row.StartRow + 1).Height + 0.1;
                    iteratedNr++;
                }

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void IteratingAndAddingToRowsWsNoValues()
        {
            using (var p = OpenPackage("IteratingAndAddingToRowsWs.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("iterateSetWs");

                var range = ws.Cells["N1:R5"];
                range.Value = "1";

                int numIter = 0;

                List<int> iteratedRows = new();

                foreach (var row in ws.Rows)
                {
                    iteratedRows.Add(row.StartRow);
                    row.Style.Fill.SetBackground(Color.CornflowerBlue);
                    ws.InsertRow(row.StartRow + 1, 1);
                    numIter++;
                }

                Assert.AreEqual(5, numIter);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void HoleInRows()
        {
            using (var p = OpenPackage("HoleInRows.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("iterateSeveral");

                var range = ws.Cells["N1:R5"];

                var range2 = ws.Cells["B8:R9"];

                var range3 = ws.Cells["C10:L12"];


                range.Value = 1;
                range2.Value = 2;
                range3.Value = 3;

                List<int> iteratedRows = new();

                foreach (var row in ws.Rows)
                {
                    iteratedRows.Add(row.StartRow);

                    if (row.StartRow < 6)
                    {
                        row.Style.Fill.SetBackground(Color.IndianRed);
                    }
                    else if (row.StartRow >= 6 && row.StartRow < 10)
                    {
                        row.Style.Fill.SetBackground(Color.DarkOliveGreen);
                    }
                    else if (row.StartRow > 9)
                    {
                        row.Style.Fill.SetBackground(Color.CadetBlue);
                    }
                }

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void RowLargerThanMaxRowShouldBeFalse()
        {
            using (var p = OpenPackage("IteratingOverRow.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("iterateSetWs");

                var range = ws.Cells["N1:R5"];

                ws.Cells["N2"].Value = 5;

                ws.Cells["R4"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Dashed;

                ws.Row(5).Hidden = true;

                var cell = ws.Cells["S6"];
                ws.Cells["S6"].Style.Fill.SetBackground(Color.Brown);

                var cs = ws._values;

                var row = cell.Start.Row + 1;
                var col = cell.Start.Column;
                //_cs.NextCell(ref enumRow, ref enumCol, enumRow, minCol, _toRow, endColumn);
                var found = cs.NextCell(ref row, ref col, row, 0, 6, col);

                Assert.IsFalse(found);
            }
        }
    }
}
