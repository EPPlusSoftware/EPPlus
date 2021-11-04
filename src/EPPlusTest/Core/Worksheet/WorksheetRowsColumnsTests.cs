using EPPlusTest;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Drawing;

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
            foreach(var row in ws.Rows[2,10])
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
                if(row.StartRow!=2 && row.StartRow!=11)
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

            ws.Cells["A1:K1"].FillNumber(x=>
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
                if(column.StartColumn < 3 || column.StartColumn > 7)
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

            Assert.AreEqual(valueCell, ws.Cells[1,2].Value);
            Assert.AreEqual(valueCell, columns.Range.GetCellValue<string>(0, 0));
            Assert.AreEqual(Style.ExcelFillStyle.LightTrellis, ws.Cells[50, 3].Style.Fill.PatternType);
            Assert.AreEqual(Color.Aqua.ToArgb().ToString("X"), ws.Cells[50, 3].Style.Fill.BackgroundColor.Rgb);
        }
    }
}