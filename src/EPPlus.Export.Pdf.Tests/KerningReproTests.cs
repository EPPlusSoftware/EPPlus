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
  09/07/2026         EPPlus Software AB           Removed the per-character reference row
 *************************************************************************************************/
using EPPlus.Export.Pdf.Tests;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.IO;
using System.Text;

namespace EPPlus.Export.Pdf.Tests
{
    /// <summary>
    /// Reproduces the reference sheet the kerning and mark-to-base work is measured against.
    ///
    /// This is a VISUAL repro, not an automated regression test. It produces a PDF that has to be
    /// opened and measured by hand, because the adjustments live inside a Flate compressed
    /// content stream and there is no decompression helper in the test project yet.
    ///
    /// What to measure in the output, at Calibri 48 pt:
    ///
    ///   A2  "AVATAR Wa To"  The kerned string. The pairs that carry kerning are A+V, V+A, A+T,
    ///                       T+A, W+a and T+o. There are TWO A-to-A pen steps in this string,
    ///                       over different pairs and with different values - measure the first
    ///                       one (over A+V) or it will not match the reference numbers.
    ///
    ///                       Unkerned, the string's pen width is 12634 font units; kerned it is
    ///                       11876. Those come from hmtx plus the kern values in the embedded
    ///                       subset, not from this sheet - there is deliberately no unkerned
    ///                       reference row. One character per cell does NOT give one: each cell
    ///                       starts a new text run at the cell origin, so the spacing would be
    ///                       the column width rather than the glyph advance.
    ///
    ///   A4  "A" + U+0301    The mark-to-base cells. Column A is DECOMPOSED (base glyph plus
    ///   A5  "a" + U+0301    combining acute) and goes through mark positioning. Column B is the
    ///   B4  precomposed A   PRECOMPOSED single glyph, which resolves in cmap and never reaches
    ///   B5  precomposed a   the mark code, so it carries the font's own accent placement.
    ///
    ///                       Column B is therefore the control: once mark-to-base offsets reach
    ///                       the renderer, the accent in A4 must line up with B4 to within a
    ///                       pixel. Before that, A4's accent sits 287 font units lower.
    ///
    ///                       A5 is expected to end up roughly 44 font units HIGHER than B5. That
    ///                       is not a defect - the anchor and the precomposed glyph disagree
    ///                       slightly, and the anchor is what shaping must follow.
    /// </summary>
    [TestClass]
    public class KerningReproTests : PdfTestBase
    {
        /// <summary>
        /// Calibri is a system font, so this test is only meaningful on a machine that has it.
        /// It is the font the original measurements were taken with; its kern lookups are
        /// extension wrapped and it has a type 4 mark lookup for the combining acute.
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
            sheet.Cells["A1"].Value = "Kerned pen steps, measure the first A to A step";
            StyleHeader(sheet.Cells["A1"]);

            sheet.Cells["A2"].Value = KernedText;
            StyleReference(sheet.Cells["A2"]);

            sheet.Cells["A3"].Value = "Mark to base. Column A decomposed, column B precomposed control";
            StyleHeader(sheet.Cells["A3"]);

            // Decomposed: base glyph plus combining acute accent, U+0301. Goes through
            // MarkToBaseProvider.
            sheet.Cells["A4"].Value = "A\u0301";
            sheet.Cells["A5"].Value = "a\u0301";

            // Precomposed, a single glyph in cmap. Carries the font's own accent placement and
            // is the control the decomposed cells must line up with.
            sheet.Cells["B4"].Value = "\u00C1";
            sheet.Cells["B5"].Value = "\u00E1";

            StyleReference(sheet.Cells["A4:B5"]);
            sheet.Cells.AutoFitColumns();

            for (int row = 1; row <= 5; row++)
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