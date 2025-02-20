using EPPlusTest;
using EPPlusTest.LoadFunctions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace OfficeOpenXml.Core.Worksheet
{
    [TestClass]
    public class WorksheetRowsColumnsTests : TestBase
    {
        static ExcelPackage _pck;

        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("WorksheetRowCol.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void ValidateRowsCollectionEnumeration()
        {
            var ws = _pck.Workbook.Worksheets.Add("Rows");

            ws.Cells["A1:A10"].FillNumber(1);

            int r = 2;
            foreach (var row in ws.Rows[2, 10])
            {
                Assert.AreEqual(r++, row.StartRow);
            }
            Assert.AreEqual(11, r);
        }
        [TestMethod]
        public void ValidateRowsCollectionEnumerationEveryOther()
        {
            var ws = _pck.Workbook.Worksheets.Add("RowsEveryOther");

            ws.Cells["A2"].Value = 2;
            ws.Cells["A4"].Value = 4;
            ws.Cells["A6"].Value = 6;
            ws.Cells["A8"].Value = 8;
            ws.Cells["A10"].Value = 10;
            int r = 2;

            foreach (var row in ws.Rows[1, 10])
            {
                Assert.AreEqual(r, row.StartRow);
                r += 2;
            }
            Assert.AreEqual(12, r);
        }
        [TestMethod]
        public void ValidateRowsCollectionEnumerationNoRows()
        {
            var ws = _pck.Workbook.Worksheets.Add("NoRows");

            ws.Cells["A1"].Value = 1;
            ws.Cells["A11"].Value = 11;

            foreach (var row in ws.Rows[2, 10])
            {
                Assert.Fail("No rows should be in the Rows collection.");
            }
        }
        [TestMethod]
        public void ValidateRowsCollectionEnumerationNoIndexerParams()
        {
            var ws = _pck.Workbook.Worksheets.Add("RowsNoIndexerParams");

            ws.Cells["A2"].Value = 2;
            ws.Cells["A11"].Value = 11;
            var rows = 0;
            foreach (var row in ws.Rows)
            {
                if (row.StartRow != 2 && row.StartRow != 11)
                {
                    Assert.Fail("Unknown row in enumeration");
                }
                rows++;
            }
            Assert.AreEqual(2, rows);
        }
        [TestMethod]
        public void ValidateColumnsCollectionEnumeration()
        {
            var ws = _pck.Workbook.Worksheets.Add("Columns");

            ws.Cells["A1:K1"].FillNumber(x =>
            {
                x.StartValue = 1;
                x.StepValue = 1;
                x.Direction = eFillDirection.Row;
            });

            int c = 2;
            foreach (var column in ws.Columns[2, 10])
            {
                Assert.AreEqual(c++, column.StartColumn);
            }
            Assert.AreEqual(11, c);
        }
        [TestMethod]
        public void ValidateColumnsCollectionEnumerationColumn3_7()
        {
            var ws = _pck.Workbook.Worksheets.Add("Columns3_7");

            ws.Columns[3, 5].Width = 25;
            ws.Cells["F3"].Value = "Column F";
            ws.Columns[7].Width = 20;

            int columns = 0;
            foreach (var column in ws.Columns[2, 10])
            {
                if (column.StartColumn < 3 || column.StartColumn > 7)
                {
                    Assert.Fail("Invalid columns detected in [Columns] collection");
                }
                columns++;
            }
            Assert.AreEqual(5, columns);
        }
        [TestMethod]
        public void ValidateColumnsCollectionEnumerationColumnWithGap()
        {
            var ws = _pck.Workbook.Worksheets.Add("ColumnsWithGap");

            ws.Columns[3].Width = 25;
            ws.Columns[8].PageBreak = true;

            ws.Cells["F3"].Value = "Column F";

            ws.Cells["J13"].Formula = "A1";
            int columns = 0;
            foreach (var column in ws.Columns[2, 10])
            {
                if (!(column.StartColumn == 3 || column.StartColumn == 8 || column.StartColumn == 6 || column.StartColumn == 10))
                {
                    Assert.Fail("Invalid columns detected in [Columns] collection");
                }

                columns++;
            }
            Assert.AreEqual(4, columns);
        }
        [TestMethod]
        public void ValidateColumnsRange()
        {
            var ws = _pck.Workbook.Worksheets.Add("ColumnsRangeProperties");

            var valueCell = "First Cell";
            var columns = ws.Columns[2, 4];
            columns.Range.SetCellValue(0, 0, valueCell);
            columns.Range.Style.Fill.SetBackground(Color.Aqua, Style.ExcelFillStyle.LightTrellis);

            Assert.AreEqual(valueCell, ws.Cells[1, 2].Value);
            Assert.AreEqual(valueCell, columns.Range.GetCellValue<string>(0, 0));
            Assert.AreEqual(Style.ExcelFillStyle.LightTrellis, ws.Cells[50, 3].Style.Fill.PatternType);
            Assert.AreEqual(Color.Aqua.ToArgb().ToString("X"), ws.Cells[50, 3].Style.Fill.BackgroundColor.Rgb);
        }
        [TestMethod]
        public void ColsDeleteAllWorks()
        {
            using (var p = OpenPackage("Cols_DeleteAllWorks.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = p.Workbook.Worksheets.Add("SomeWs");

                var table = ws.Tables.Add(ws.Cells["A1:K5"], "ATable");
                table.Columns[0].Name = "SomeName";
                table.Columns[1].Name = "SomeBODY";
                table.Columns[4].Name = "OnceTold";
                table.Columns[5].Name = "Me";
                table.Columns[6].Name = "SomeWorld";

                table.SyncColumnNames(Table.ApplyDataFrom.ColumnNamesToCells);
                ws.Columns[1].Hidden = true;
                ws.Columns[2].Hidden = true;
                ws.Columns[7].Hidden = true;
                ws.Columns[4].AutoFit();
                ws.Columns[5].AutoFit();
                ws.Columns[6].AutoFit();
                ws.Columns[7].AutoFit();

                ws.Cells["A2:K2"].Formula = "COLUMN()";
                ws.Calculate();
                ws.Cells.ClearFormulas();

                var val = ws.Cells["A1"].Value;

                ws.Columns.DeleteAll(col => col.Hidden);

                table.SyncColumnNames(Table.ApplyDataFrom.ColumnNamesToCells);

                Assert.AreEqual("Column3", ws.Cells["A1"].Value);
                Assert.AreEqual("Column4", ws.Cells["B1"].Value);
                Assert.AreEqual("OnceTold", ws.Cells["C1"].Value);
                Assert.AreEqual("Me", ws.Cells["D1"].Value);
                Assert.AreEqual("Column8", ws.Cells["E1"].Value);

                Assert.AreEqual(8, table.Columns.Count());

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void ColsIterateAndDeleteWithPredicate()
        {
            using (var p = OpenPackage("Cols_IterateAndDeleteWithPredicate.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = p.Workbook.Worksheets.Add("SomeWs");

                var range = ws.Cells["N1:R5"];
                range.Value = "";

                var cols = range.EntireColumn;

                foreach (var col in cols)
                {
                    ws.Columns[col.StartColumn].Style.Fill.SetBackground(Color.Red);
                }

                var subRange = ws.Cells["O1:Q5"];
                var subColRange = subRange.EntireColumn;

                List<int> iteratedColumns = new();

                foreach (var column in subColRange)
                {
                    ws.Columns[column.StartColumn].Style.Fill.SetBackground(Color.Blue);
                    iteratedColumns.Add(column.StartColumn);
                }

                //Ensure columns do not iterate beyond ToCol
                Assert.AreEqual(3, iteratedColumns.Count());
                Assert.AreEqual(17, iteratedColumns[2]);

                var newRange = ws.Cells["M1:S5"];

                var rgb = ws.Cells["N1"].Style.Fill.BackgroundColor.Rgb;
                ws.Columns.DeleteAll(column => column.Style.Fill.BackgroundColor.Rgb == rgb);

                var aCol = ws.Columns[range.Start.Column];

                var secondToLastCol = ws.Columns[range.End.Column - 1].Style.Fill.BackgroundColor.Rgb;
                var lastCol = ws.Columns[range.End.Column].Style.Fill.BackgroundColor.Rgb;

                //Ensure after two columns deleted the last two columns updated
                Assert.IsNull(lastCol);
                Assert.IsNull(secondToLastCol);

                foreach (var col in ws.Columns[range.Start.Column, range.Start.Column + 2])
                {
                    //Ensure all three columns in start of range are now blue
                    Assert.AreEqual("FF0000FF", col.Style.Fill.BackgroundColor.Rgb);
                }

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void ColsDeleteMiddle()
        {
            using (var p = OpenPackage("Cols_DeleteMiddle.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = p.Workbook.Worksheets.Add("SomeWs");

                var range = ws.Cells["N1:R5"];
                range.Value = "";

                var cols = ws.Columns[range.Start.Column, range.End.Column];
                var colsByRange = range.EntireColumn;

                foreach (var col in cols)
                {
                    ws.Columns[col.StartColumn].Style.Fill.SetBackground(Color.Yellow);
                }

                var subRange = ws.Cells["O1:Q5"];

                var subColRange = ws.Columns[subRange.Start.Column, subRange.End.Column - 1];

                for (int i = subRange.Start.Column; i <= subRange.End.Column; i++)
                {
                    ws.Columns[i].Style.Fill.SetBackground(Color.Green);
                }

                var greenRgb = ws.Cells["O1"].Style.Fill.BackgroundColor.Rgb;
                var yellowRgb = ws.Columns[range.Start.Column].Style.Fill.BackgroundColor.Rgb;

                ws.Columns.DeleteAll(column => column.Style.Fill.BackgroundColor.Rgb == greenRgb);

                var remaingingCols = ws.Columns[range.Start.Column, range.Start.Column + 1];
                foreach (var col in remaingingCols)
                {
                    //Ensure the two remaining columns are now yellow
                    Assert.AreEqual(yellowRgb, col.Style.Fill.BackgroundColor.Rgb);
                }

                p.SaveAs(GetOutputFile("", "ColAfterDelete.xlsx").FullName);
            }
        }

        [TestMethod]
        public void ColsNotSetWhenIteratingEmptyCS()
        {
            using (var p = OpenPackage("Cols_NotSetWhenIteratingEmptyCS.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = p.Workbook.Worksheets.Add("SomeWs");

                var range = ws.Cells["N1:R5"];
                //Since no value exists cellstore is not set
                //range.Value = "";

                var cols = ws.Columns[range.Start.Column, range.End.Column];

                foreach (var col in cols)
                {
                    ws.Columns[col.StartColumn].Style.Fill.SetBackground(Color.Red);
                }

                var rgbColumn = ws.Columns[range.Start.Column].Style.Fill.BackgroundColor.Rgb;
                var rgbCell = ws.Cells["N1"].Style.Fill.BackgroundColor.Rgb;

                Assert.IsNull(rgbColumn);
                Assert.IsNull(rgbCell);

                Assert.AreEqual(rgbColumn, rgbCell);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void ColsSetColorOnRangeShouldWorkWhenEmptyCs()
        {
            using (var p = OpenPackage("Cols_SetColorOnRangeShouldWorkWhenEmptyCs.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = p.Workbook.Worksheets.Add("SomeWs");

                var range = ws.Cells["N1:R5"];

                range.EntireColumn.Style.Fill.SetBackground(Color.Red);

                var rgbColumn = ws.Columns[range.Start.Column].Style.Fill.BackgroundColor.Rgb;
                var rgbCell = ws.Cells["N1"].Style.Fill.BackgroundColor.Rgb;

                Assert.IsNotNull(rgbColumn);
                Assert.IsNotNull(rgbCell);

                Assert.AreEqual(rgbColumn, rgbCell);

                SaveAndCleanup(p);
            }
        }
    }
}