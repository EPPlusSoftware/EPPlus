using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest
{
    [TestClass]
    public class AutofitTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("Worksheet.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;

            SaveAndCleanup(_pck);
            if (File.Exists(fileName))
            {
                File.Copy(fileName, dirName + "\\WorksheetRead.xlsx", true);
            }
        }

        [TestMethod]
        public void AutoFitColumns()
        {
            var ws = _pck.Workbook.Worksheets.Add("Autofit");
            ws.Cells["A1:H1"].Value = "Auto fit column that is veeery long...";
            ws.Cells["A1:H1"].Style.Font.Name = "Arial";
            ws.Cells["B1"].Style.TextRotation = 30;
            ws.Cells["C1"].Style.TextRotation = 45;
            ws.Cells["D1"].Style.TextRotation = 75;
            ws.Cells["E1"].Style.TextRotation = 90;
            ws.Cells["F1"].Style.TextRotation = 120;
            ws.Cells["G1"].Style.TextRotation = 135;
            ws.Cells["H1"].Style.TextRotation = 180;
            ws.Cells["A1:H1"].AutoFitColumns(0);

            ws.Column(40).AutoFit();
        }
        [TestMethod]
        public void AutoFitColumn()
        {
            var ws = _pck.Workbook.Worksheets.Add("Autofit2");
            ws.Cells["A1:A10"].Value = "Auto fit column that is veeery long...";
            ws.Cells["A1:A10"].Style.Font.Name = "Arial";
            ws.Columns[1].AutoFit();
        }

        [TestMethod]
        public void AutoFitColumnTest()
        {
            var p = OpenTemplatePackage("AutoFitWorkbook.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var start = DateTime.Now;
            ws.Columns[1].AutoFit();
            var end = DateTime.Now;
            TimeSpan span = end - start;
            Assert.AreEqual(125d, ws.Columns[1].Width, 5d);
            SaveAndCleanup(p);
        }

        [TestMethod]
        public void AutofitAutofilterTest()
        {
            using var package = OpenTemplatePackage("AutoFitAutofilter.xlsx");
            var ws = package.Workbook.Worksheets.Add("Sheet1");

            // Headers are the widest text in each column - the data below is deliberately
            // shorter so the column width is driven by the header + the autofilter dropdown arrow.
            ws.Cells["A1"].Value = "Department";
            ws.Cells["B1"].Value = "Annual Budget";
            ws.Cells["C1"].Value = "Region Name";

            // Data rows - all shorter than the headers above them.
            ws.Cells["A2"].Value = "Sales";
            ws.Cells["B2"].Value = 1200;
            ws.Cells["C2"].Value = "North";

            ws.Cells["A3"].Value = "IT";
            ws.Cells["B3"].Value = 980;
            ws.Cells["C3"].Value = "West";

            ws.Cells["A4"].Value = "HR";
            ws.Cells["B4"].Value = 540;
            ws.Cells["C4"].Value = "East";

            // Apply autofilter across the header row + data.
            ws.Cells["A1:C4"].AutoFilter = true;

            // Autofit the columns.
            ws.Cells["A1:C4"].AutoFitColumns();

            // Inspect what EPPlus actually produced for each column.
            System.Diagnostics.Debug.WriteLine($"Column A (Department):    {ws.Column(1).Width}");
            System.Diagnostics.Debug.WriteLine($"Column B (Annual Budget): {ws.Column(2).Width}");
            System.Diagnostics.Debug.WriteLine($"Column C (Region Name):   {ws.Column(3).Width}");

            // Save the workbook
            SaveAndCleanup(package);
        }

        [TestMethod]
        public void AutoFitColumnsWithAutoFilter()
        {
            var ws = _pck.Workbook.Worksheets.Add("AutofitAutoFilter");
            ws.Cells["A1"].Value = "hour";
            ws.Cells["B1"].Value = "minute";
            ws.Cells["A2"].Value = 12;
            ws.Cells["B2"].Value = 30;

            ws.Cells["A1:B2"].AutoFilter = true;

            ws.Cells["A1:B2"].AutoFitColumns();

            // Without the fix, the AutoFilter header row range (A1:B1) is measured as a whole.
            // Under the hood, worksheet.Cells["A1:B1"].TextForWidth evaluated to "System.Object[,]" (16 chars),
            // which forced a minimum width of ~16.07 points.
            // With the fix, the specific cell for each column in the AutoFilter is measured, 
            // resulting in a narrow width matching "hour" / "minute".
            Assert.IsTrue(ws.Column(1).Width < 12d, $"Column 1 width should be small but was {ws.Column(1).Width}");
            Assert.IsTrue(ws.Column(2).Width < 12d, $"Column 2 width should be small but was {ws.Column(2).Width}");
        }

        [TestMethod]
        public void Autofit_Skip()
        {
            // Skip: a WrapText cell must not contribute to the column width at all.
            // Column A holds a long wrapped cell; column B holds nothing. Under Skip the
            // wrapped cell is ignored, so both columns end up at the same (default) width.
            using var package = new ExcelPackage();
            package.Settings.TextSettings.WrappedTextAutofitMode = eWrappedTextAutofitMode.Skip;
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = "Line 1\nLine 2 is a bit longer\nLine 3";
            ws.Cells["A1"].Style.WrapText = true;
            // Act
            ws.Cells["A1:B1"].AutoFitColumns();
            // Assert
            Assert.AreEqual(ws.Columns[2].Width, ws.Columns[1].Width, 0.0001d,
                "Skip should ignore the wrapped cell, so column A matches the empty column B");
        }

        [TestMethod]
        public void Autofit_FullText()
        {
            // FullText: the entire cell text is measured as a single line.
            // The reference cell B holds the same text; because a WrapText cell keeps its
            // newlines (measured as zero width), B must use identical newline placement so
            // both cells measure the exact same visible characters.
            using var package = new ExcelPackage();
            package.Settings.TextSettings.WrappedTextAutofitMode = eWrappedTextAutofitMode.FullText;
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = "Line 1\nLine 2 is a bit longer\nLine 3";
            ws.Cells["B1"].Value = "Line 1\nLine 2 is a bit longer\nLine 3";
            ws.Cells["A1:B1"].Style.WrapText = true;
            // Act
            ws.Cells.AutoFitColumns();
            // Assert
            Assert.AreEqual(ws.Columns[2].Width, ws.Columns[1].Width, 0.0001d,
                "FullText should measure the entire string, matching the identical reference cell");
        }

        [TestMethod]
        public void Autofit_SplitNewLine()
        {
            // SplitNewLine: the widest newline-separated line drives the width.
            // Reference cell B holds that line ("Line 2 is a bit longer") on its own.
            using var package = new ExcelPackage();
            package.Settings.TextSettings.WrappedTextAutofitMode = eWrappedTextAutofitMode.SplitNewLine;
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = "Line 1\nLine 2 is a bit longer\nLine 3";
            ws.Cells["B1"].Value = "Line 2 is a bit longer";
            ws.Cells["A1:B1"].Style.WrapText = true;
            // Act
            ws.Cells.AutoFitColumns();
            // Assert
            Assert.AreEqual(ws.Columns[2].Width, ws.Columns[1].Width, 0.0001d,
                "Column A (widest line) should match column B (that line in full)");
        }

        [TestMethod]
        public void Autofit_SplitWord()
        {
            // SplitWord: the widest whitespace/hyphen-separated segment drives the width.
            // Reference cell B holds that word ("aVeryLongWord") on its own.
            using var package = new ExcelPackage();
            package.Settings.TextSettings.WrappedTextAutofitMode = eWrappedTextAutofitMode.SplitWord;
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = "short longer aVeryLongWord medium";
            ws.Cells["B1"].Value = "aVeryLongWord";
            ws.Cells["A1:B1"].Style.WrapText = true;
            // Act
            ws.Cells.AutoFitColumns();
            // Assert
            Assert.AreEqual(ws.Columns[2].Width, ws.Columns[1].Width, 0.0001d,
                "Column A (widest word) should match column B (that word in full)");
        }

        [TestMethod]
        public void Autofit_SplitWord_HyphenIsVisibleAndBreaksTheWord()
        {
            // A hyphen is a visible break boundary: it terminates the preceding segment
            // AND its own width is counted into that segment. So the widest segment of
            // "aVeryLongWord-x" is "aVeryLongWord-" (word + trailing hyphen), not the bare word.
            using var package = new ExcelPackage();
            package.Settings.TextSettings.WrappedTextAutofitMode = eWrappedTextAutofitMode.SplitWord;
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = "aVeryLongWord-x";
            ws.Cells["B1"].Value = "aVeryLongWord-";   // includes the trailing hyphen
            ws.Cells["C1"].Value = "aVeryLongWord";    // bare word, no hyphen
            ws.Cells["A1:C1"].Style.WrapText = true;
            // Act
            ws.Cells.AutoFitColumns();
            // Assert
            Assert.AreEqual(ws.Columns[2].Width, ws.Columns[1].Width, 0.0001d,
                "Segment should include the trailing hyphen (visible boundary)");
            Assert.IsTrue(ws.Columns[1].Width > ws.Columns[3].Width,
                "Segment with hyphen should be wider than the bare word without it");
        }

        [TestMethod]
        public void Autofit_SplitNewLine_CrlfCountsAsSingleBreak()
        {
            // A CRLF pair must be treated as one line break, not two. If it were counted
            // as two breaks it would create an empty phantom segment between the lines,
            // but that has zero width and would not change the result. What this guards
            // is that the CR is not measured as a separate one-character segment and that
            // the widest line is identified correctly across a CRLF boundary.
            using var package = new ExcelPackage();
            package.Settings.TextSettings.WrappedTextAutofitMode = eWrappedTextAutofitMode.SplitNewLine;
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = "Short\r\nLine 2 is a bit longer\r\nShort";
            ws.Cells["B1"].Value = "Line 2 is a bit longer";
            ws.Cells["A1:B1"].Style.WrapText = true;
            // Act
            ws.Cells.AutoFitColumns();
            // Assert
            Assert.AreEqual(ws.Columns[2].Width, ws.Columns[1].Width, 0.0001d,
                "CRLF should be treated as a single line break; widest line should match column B");
        }

        [TestMethod]
        public void Autofit_SplitNewLine_WidestSegmentFirstIsStillChosen()
        {
            // The widest segment appears FIRST here. This guards against a reset bug where
            // the running width of an earlier (wider) segment fails to carry over into the
            // max comparison once a later, narrower segment is measured.
            using var package = new ExcelPackage();
            package.Settings.TextSettings.WrappedTextAutofitMode = eWrappedTextAutofitMode.SplitNewLine;
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = "Line 1 is clearly the longest\nShort\nAlso short";
            ws.Cells["B1"].Value = "Line 1 is clearly the longest";
            ws.Cells["A1:B1"].Style.WrapText = true;
            // Act
            ws.Cells.AutoFitColumns();
            // Assert
            Assert.AreEqual(ws.Columns[2].Width, ws.Columns[1].Width, 0.0001d,
                "The widest line is the first one and must still drive the column width");
        }

        [TestMethod]
        public void Autofit_SplitNewLine_EastAsianWidthResetsPerLine()
        {
            // Regression test for the East Asian width bug: previously the EA width (widthEA)
            // accumulated across ALL lines and was never reset at a line break, so a multi-line
            // CJK cell was measured as the SUM of every line's EA width instead of the widest
            // single line. Here the three lines are 2, 5 and 3 hiragana characters; the correct
            // result is the width of the 5-character line. With the bug it would be roughly the
            // width of all 10 characters combined.
            using var package = new ExcelPackage();
            package.Settings.TextSettings.WrappedTextAutofitMode = eWrappedTextAutofitMode.SplitNewLine;
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = "\u3042\u3042\n\u3042\u3042\u3042\u3042\u3042\n\u3042\u3042\u3042"; // ,  , 
            ws.Cells["B1"].Value = "\u3042\u3042\u3042\u3042\u3042";                                   //  (widest line)
            ws.Cells["A1:B1"].Style.WrapText = true;
            // Act
            ws.Cells.AutoFitColumns();
            // Assert
            Assert.AreEqual(ws.Columns[2].Width, ws.Columns[1].Width, 0.0001d,
                "Multi-line CJK column should match its widest line, not the sum of all lines");
        }

        [TestMethod]
        public void AutoFitTestWithDifferenLengths()
        {
            using (var p = OpenPackage("SimpleAutofitTests.xlsx", true))
            {
                //p.Settings.TextSettings.WrappedTextAutofitMode = eWrappedTextAutofitMode.SplitWord;
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = "Little";
                ws.Cells["A2"].Value = "MEDIUM";
                ws.Cells["A3"].Value = "Largeeeeeesssst";
                ws.Cells["A4"].Value = "Larg-ish";

                ws.Cells["B1"].Value = "I should not be autofit";

                var bWidth = ws.Columns[2].Width;

                ws.Cells["A1:A4"].AutoFitColumns();

                var bWidthAfter = ws.Columns[2].Width;

                //Untouched cols should remain untouched
                Assert.AreEqual(bWidth, bWidthAfter);

                ws.Cells["A5"].Value = "Very large but outside the range of what should be fitted";

                var widthBeforeA = ws.Columns[1].Width;

                ws.Cells["A1:A4"].AutoFitColumns();

                var widthAfterA = ws.Columns[1].Width;


                //Untouched cells within same column Probaly should not change the column
                //technically different from excel but also different syntax
                Assert.AreEqual(widthBeforeA, widthAfterA);


                ws.Cells.AutoFitColumns();

                //Doing all cells should however
                Assert.AreNotEqual(widthAfterA, ws.Columns[1].Width);
                Assert.IsTrue(ws.Columns[1].Width > widthAfterA);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void AutofitOneCellCompoundingConfigs()
        {
            using (var p = OpenPackage("AutofitCompoundOneCell.xlsx", true))
            {
                //p.Settings.TextSettings.WrappedTextAutofitMode = eWrappedTextAutofitMode.SplitWord;
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = "aaaaaaaaaaaaaaaaa";

                ws.Cells["A1"].AutoFitColumns();

                ws.Cells["A1"].Value = "aaaaaaaaaaaaaaaaaaaaaaaa";
                ws.Cells["A1"].Style.WrapText = true;

                ws.Cells["A1"].AutoFitColumns();

                var colWidth = ws.Column(1).Width;;

                SaveAndCleanup(p);

                //Does not appear to match output file
                //Might still be correct bc OS margins etc.
                Assert.AreEqual(8.43d, ws.Column(1).Width);
            }
        }
    }
}
