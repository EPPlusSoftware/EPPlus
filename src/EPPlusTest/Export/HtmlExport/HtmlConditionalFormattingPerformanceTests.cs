using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using System;
using System.Diagnostics;
using System.Drawing;

namespace EPPlusTest.Export.HtmlExport
{
    /// <summary>
    /// Conditional formatting rules that compare a cell against the whole range must not re-evaluate the range for every cell,
    /// and must not walk the empty cells of a range that covers a full sheet.
    /// </summary>
    [TestClass]
    public class HtmlConditionalFormattingPerformanceTests : TestBase
    {
        const int LargeRowCount = 50000;
        static readonly TimeSpan MaxExportTime = TimeSpan.FromSeconds(30);

        private static ExcelWorksheet CreateSheet(ExcelPackage p, int rows)
        {
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            for (int r = 1; r <= rows; r++)
            {
                ws.Cells[r, 1].Value = r % 7;
            }
            return ws;
        }

        private static void AssertExportIsFast(ExcelRangeBase range)
        {
            var sw = Stopwatch.StartNew();
            var html = range.CreateHtmlExporter().GetHtmlString();
            sw.Stop();
            Assert.IsFalse(string.IsNullOrEmpty(html));
            Assert.IsTrue(sw.Elapsed < MaxExportTime, $"Html export took {sw.Elapsed}.");
        }

        [TestMethod]
        public void DuplicateValuesOnLargeRangeShouldExportInLinearTime()
        {
            using (var p = new ExcelPackage())
            {
                var ws = CreateSheet(p, LargeRowCount);
                var range = ws.Cells[1, 1, LargeRowCount, 1];
                range.ConditionalFormatting.AddDuplicateValues();
                AssertExportIsFast(range);
            }
        }

        [TestMethod]
        public void UniqueValuesOnLargeRangeShouldExportInLinearTime()
        {
            using (var p = new ExcelPackage())
            {
                var ws = CreateSheet(p, LargeRowCount);
                var range = ws.Cells[1, 1, LargeRowCount, 1];
                range.ConditionalFormatting.AddUniqueValues();
                AssertExportIsFast(range);
            }
        }

        [TestMethod]
        public void TopBottomOnLargeRangeShouldExportInLinearTime()
        {
            using (var p = new ExcelPackage())
            {
                var ws = CreateSheet(p, LargeRowCount);
                var range = ws.Cells[1, 1, LargeRowCount, 1];
                range.ConditionalFormatting.AddTop();
                range.ConditionalFormatting.AddTopPercent();
                range.ConditionalFormatting.AddBottom();
                range.ConditionalFormatting.AddBottomPercent();
                AssertExportIsFast(range);
            }
        }

        [TestMethod]
        public void ColorScalesOnLargeRangeShouldExportInLinearTime()
        {
            using (var p = new ExcelPackage())
            {
                var ws = CreateSheet(p, LargeRowCount);
                var range = ws.Cells[1, 1, LargeRowCount, 1];
                range.ConditionalFormatting.AddTwoColorScale();
                range.ConditionalFormatting.AddThreeColorScale();
                AssertExportIsFast(range);
            }
        }

        [TestMethod]
        public void StdDevOnLargeRangeShouldExportInLinearTime()
        {
            using (var p = new ExcelPackage())
            {
                var ws = CreateSheet(p, LargeRowCount);
                var range = ws.Cells[1, 1, LargeRowCount, 1];
                range.ConditionalFormatting.AddAboveStdDev();
                range.ConditionalFormatting.AddBelowStdDev();
                AssertExportIsFast(range);
            }
        }

        [TestMethod]
        public void PercentileIconSetAndDataBarOnLargeRangeShouldExportInLinearTime()
        {
            using (var p = new ExcelPackage())
            {
                var ws = CreateSheet(p, LargeRowCount);
                var range = ws.Cells[1, 1, LargeRowCount, 1];
                var iconSet = range.ConditionalFormatting.AddThreeIconSet(eExcelconditionalFormatting3IconsSetType.Arrows);
                iconSet.Icon2.Type = eExcelConditionalFormattingValueObjectType.Percentile;
                iconSet.Icon2.Value = 33;
                iconSet.Icon3.Type = eExcelConditionalFormattingValueObjectType.Percentile;
                iconSet.Icon3.Value = 67;

                var dataBar = range.ConditionalFormatting.AddDatabar(Color.Blue);
                dataBar.LowValue.Type = eExcelConditionalFormattingValueObjectType.Percentile;
                dataBar.LowValue.Value = 10;
                dataBar.HighValue.Type = eExcelConditionalFormattingValueObjectType.Percentile;
                dataBar.HighValue.Value = 90;
                AssertExportIsFast(range);
            }
        }

        [TestMethod]
        public void RulesOverFullSheetShouldNotWalkEmptyCells()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = 1;
                ws.Cells["A2"].Value = 1;
                ws.Cells["B3"].Value = 5;
                const string fullSheet = "A1:XFD1048576";
                ws.ConditionalFormatting.AddDuplicateValues(fullSheet);
                ws.ConditionalFormatting.AddUniqueValues(fullSheet);
                ws.ConditionalFormatting.AddTop(fullSheet);
                ws.ConditionalFormatting.AddTwoColorScale(fullSheet);
                ws.ConditionalFormatting.AddThreeColorScale(fullSheet);
                ws.ConditionalFormatting.AddThreeIconSet(fullSheet, eExcelconditionalFormatting3IconsSetType.Arrows);
                ws.ConditionalFormatting.AddDatabar(fullSheet, Color.Blue);

                AssertExportIsFast(ws.Cells["A1:B3"]);
            }
        }

        [TestMethod]
        public void DuplicateAndUniqueValuesShouldMatchExpectedCells()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = "a";
                ws.Cells["A2"].Value = "A";
                ws.Cells["A3"].Value = "b";
                ws.Cells["A4"].Value = 1;
                ws.Cells["A5"].Value = 2;
                ws.Cells["A6"].Value = 2;

                var dup = (ExcelConditionalFormattingDuplicateValues)ws.ConditionalFormatting.AddDuplicateValues("A1:A6");
                var uni = (ExcelConditionalFormattingUniqueValues)ws.ConditionalFormatting.AddUniqueValues("A1:A6");

                var expectedDup = new[] { true, true, false, false, true, true };
                var expectedUnique = new[] { true, true, true, true, false, false }; //Unique compares the values case sensitive.
                for (int r = 1; r <= 6; r++)
                {
                    Assert.AreEqual(expectedDup[r - 1], dup.ShouldApplyToCell(new ExcelAddress(r, 1, r, 1)), $"Duplicate A{r}");
                    Assert.AreEqual(expectedUnique[r - 1], uni.ShouldApplyToCell(new ExcelAddress(r, 1, r, 1)), $"Unique A{r}");
                }
                ws.ConditionalFormatting.ClearTempExportCacheForAllCFs();
            }
        }

        [TestMethod]
        public void CacheShouldBeClearedBetweenExports()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = 1;
                ws.Cells["A2"].Value = 2;
                ws.Cells["A3"].Value = 3;
                var range = ws.Cells["A1:A3"];
                range.ConditionalFormatting.AddDuplicateValues().Style.Fill.BackgroundColor.SetColor(Color.Red);
                var first = range.CreateHtmlExporter().GetHtmlString();

                ws.Cells["A3"].Value = 2;
                var second = range.CreateHtmlExporter().GetHtmlString();

                Assert.AreNotEqual(first, second);
            }
        }
    }
}
