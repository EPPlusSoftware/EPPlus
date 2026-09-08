/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/07/2026         EPPlus Software AB           WP1 kerning repro sheet
 *************************************************************************************************/
using EPPlus.Export.Pdf.Tests;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.IO;
using System.Text;

namespace EPPlusTest.PDF
{
    /// <summary>
    /// Reproduces the reference sheet the WP1 kerning work was measured against.
    ///
    /// This is a VISUAL repro, not an automated regression test. It produces a PDF that has to be
    /// opened and measured by hand, because the kerning adjustments live inside a Flate compressed
    /// content stream and there is no decompression helper in the test project yet.
    ///
    /// What to measure in the output, at Calibri 48 pt and 3.3299 px/pt:
    ///
    ///   A2  "AVATAR Wa To"  The A to A pen step is the reference measurement. It was 184 px
    ///                       before WP1 and should now be close to 170 px, which is 2358 font
    ///                       units down to 2184. The pairs that carry kerning here are A+V, V+A,
    ///                       A+T, T+A, W+a and T+o.
    ///
    ///   A4  "AVATAR Wa To"  Same string with kerning suppressed by putting every character in its
    ///   A5                  own cell, as a side by side reference for the pen steps. Not affected
    ///                       by WP1.
    ///
    ///   A7  "A" + U+0301    The accent cells are expected to be UNCHANGED by WP1. XOffset and
    ///   A8  "a" + U+0301    YOffset are still not read by the renderer, so the accent over the
    ///                       capital A should still sit about 7 px too far right and 23 px too low
    ///                       and collide with the apex, while the accent over the lower case a
    ///                       still looks correct by coincidence. Both moving is a sign that WP1
    ///                       touched something it should not have. That is WP2.
    ///
    /// The decomposed sequences in A7 and A8 are the ones that exercise mark to base positioning.
    /// The precomposed characters in B7 and B8 resolve to a single glyph in cmap and never reach
    /// the mark positioning code, so they are included only as a visual baseline.
    /// </summary>
    [TestClass]
    public class KerningReproTests : PdfTestBase
    {
        /// <summary>
        /// Calibri is a system font, so this test is only meaningful on a machine that has it.
        /// It is the font the original measurements were taken with, and its kern lookups are
        /// extension wrapped, which is the failure mode WP1 fixes.
        /// </summary>
        private const string ReferenceFontName = "Calibri";

        private const float ReferenceFontSize = 48f;

        private const string KernedText = "AVATAR Wa To";

        [TestMethod]
        public void Avatar_Calibri48_KerningReproSheet()
        {
            using (var package = OpenPackage("WP1KerningRepro.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("Kerning");

                BuildReproSheet(sheet);

                // Written to disk for visual inspection and for measuring against the reference
                // image. SaveAsPdf on the worksheet exports that single sheet.
                SaveAsPdf(sheet, "WP1KerningRepro");

                // Also export to a stream so the test fails loudly if the export itself breaks,
                // rather than silently writing an unreadable file.
                using (var stream = new MemoryStream())
                {
                    sheet.SaveAsPdf(stream);
                    AssertLooksLikePdf(stream.ToArray());
                }

                SaveWorkbook("WP1KerningRepro.xlsx", package);
            }
        }

        private static void BuildReproSheet(ExcelWorksheet sheet)
        {
            sheet.Cells["A1"].Value = "Kerned pen steps";
            StyleHeader(sheet.Cells["A1"]);

            sheet.Cells["A2"].Value = KernedText;
            StyleReference(sheet.Cells["A2"]);

            sheet.Cells["A3"].Value = "Unkerned reference, one character per cell";
            StyleHeader(sheet.Cells["A3"]);

            // One character per cell removes every pair boundary, so these advances are the raw
            // hmtx widths regardless of whether kerning works.
            for (int i = 0; i < KernedText.Length; i++)
            {
                var cell = sheet.Cells[4, i + 1];
                cell.Value = KernedText[i].ToString();
                StyleReference(cell);
            }

            sheet.Cells["A6"].Value = "Mark to base, expected unchanged by WP1";
            StyleHeader(sheet.Cells["A6"]);

            // Decomposed: base glyph plus combining acute accent, U+0301.
            sheet.Cells["A7"].Value = "A\u0301";
            sheet.Cells["A8"].Value = "a\u0301";

            // Precomposed, single glyph in cmap. Visual baseline only.
            sheet.Cells["B7"].Value = "\u00C1";
            sheet.Cells["B8"].Value = "\u00E1";

            StyleReference(sheet.Cells["A7:B8"]);

            // Wide enough that autofit or clipping never influences the measurement. Kerning that
            // starts working narrows the measured text, so a fixed width keeps the pen steps the
            // only thing that changes between runs.
            for (int column = 1; column <= KernedText.Length; column++)
            {
                sheet.Column(column).Width = 20;
            }

            for (int row = 1; row <= 8; row++)
            {
                sheet.Row(row).CustomHeight = true;
                sheet.Row(row).Height = 70;
            }
        }

        private static void StyleReference(ExcelRange range)
        {
            range.Style.Font.Name = ReferenceFontName;
            range.Style.Font.Size = ReferenceFontSize;
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

            // Left aligned and unindented, so the first glyph starts at the cell edge and the
            // A to A distance is a pure pen step with no left margin effects.
            range.Style.Indent = 0;
            range.Style.WrapText = false;
        }

        private static void StyleHeader(ExcelRange range)
        {
            range.Style.Font.Name = ReferenceFontName;
            range.Style.Font.Size = 11f;
            range.Style.Font.Bold = true;
        }

        private static void AssertLooksLikePdf(byte[] bytes)
        {
            Assert.IsTrue(bytes.Length > 0, "PDF output is empty.");

            string head = Encoding.ASCII.GetString(bytes, 0, Math.Min(8, bytes.Length));
            Assert.IsTrue(head.StartsWith("%PDF-"), $"Missing PDF header. Got: '{head}'");

            int tailLength = Math.Min(8, bytes.Length);
            string tail = Encoding.ASCII.GetString(bytes, bytes.Length - tailLength, tailLength);
            Assert.IsTrue(tail.Contains("%%EOF"), "Missing %%EOF trailer marker.");
        }
    }
}