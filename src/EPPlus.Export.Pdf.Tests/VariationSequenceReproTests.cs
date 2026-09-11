/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/09/2026         EPPlus Software AB           Unicode Variation Sequence visual repro sheet
 *************************************************************************************************/
using EPPlus.Export.Pdf.Settings;
using EPPlus.Export.Pdf.Tests;
using EPPlus.Fonts.OpenType;
using OfficeOpenXml;
using OfficeOpenXml.Export.PdfExport;
using OfficeOpenXml.Style;
using System;
using System.IO;
using System.Text;

namespace EPPlus.Export.Pdf.Tests
{
    /// <summary>
    /// VISUAL repro for cmap format 14 (Unicode Variation Sequences) actually being consulted
    /// during shaping, and preserved through subsetting and embedding - not an automated
    /// regression test, since the point is to look at the two glyph shapes with your own eyes.
    /// The three unit-test files already cover the automated side of this.
    ///
    /// Uses BIZ UDGothic, a real Japanese font bundled with the test suite that has genuine
    /// format-14 data - verified directly against the font's own bytes (not assumed): U+585A
    /// (塚) has a registered NON-DEFAULT variation sequence under VS01 (U+FE00). The base
    /// character alone resolves to glyph 3363 in this font; U+585A + VS01 resolves to glyph
    /// 1399 - a different, hand-drawn glyph, not a font-side no-op.
    ///
    /// What to look for in the exported PDF:
    ///
    ///   A2  "塚"            The plain base character - glyph 3363, the font's default form.
    ///   A3  "塚" + VS01      The variation sequence - should render glyph 1399: a VISIBLY
    ///                        different shape for the top of the right-hand component (this is
    ///                        one of the standard textbook examples of a Japanese IVS/
    ///                        hanyo-denshi variant pair).
    ///
    /// If A2 and A3 look IDENTICAL, the variation sequence is silently falling back to the base
    /// glyph somewhere in the pipeline - shaping, subsetting, or serialization.
    /// </summary>
    [TestClass]
    public class VariationSequenceReproTests : PdfTestBase
    {
        // BIZUDGothic-Regular.ttf lives in the Fonts subfolder of the test project and is
        // copied next to the test assembly at build time - no dependency on BIZ UDGothic being
        // installed as a system font.
        private static string FontsFolder => Path.Combine(AppContext.BaseDirectory, "Fonts");

        private const string ReproFontFamily = "BIZ UDGothic";
        private const float ReproFontSize = 72f;

        // U+585A (塚): base char alone -> glyph 3363, base char + VS01 (U+FE00) -> glyph 1399.
        // Both values read directly out of BIZUDGothic-Regular.ttf's own cmap tables.
        private const string BaseChar = "\u585A";
        private const string BaseCharPlusVs01 = "\u585A\uFE00";

        private static OpenTypeFontEngine CreateEngine()
        {
            return new OpenTypeFontEngine(cfg =>
            {
                cfg.FontDirectories.Add(FontsFolder);
                cfg.SearchSystemDirectories = false;
            });
        }

        [TestMethod]
        public void VariationSequence_BizUdGothic_ReproSheet()
        {
            using (var package = OpenPackage("VariationSequenceRepro.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("VariationSequence");
                BuildReproSheet(sheet);

                var engine = CreateEngine();
                var settings = new PdfPageSettings(engine);

                // Written to disk for visual inspection - this is the whole point of the test.
                SaveAsPdf(sheet, "VariationSequenceRepro", settings);

                // Also export to a stream so the test fails loudly if the export itself breaks,
                // rather than silently writing an unreadable file.
                using (var stream = new MemoryStream())
                {
                    new PdfCatalog(settings, sheet).Save(stream);
                    AssertLooksLikePdf(stream.ToArray());
                }

                SaveWorkbook("VariationSequenceRepro.xlsx", package);
            }
        }

        private static void BuildReproSheet(ExcelWorksheet sheet)
        {
            sheet.Cells["A1"].Value = "Compare A2 (plain) vs A3 (base + VS01) - look for a different top-right stroke shape";
            StyleHeader(sheet.Cells["A1"]);

            sheet.Cells["A2"].Value = BaseChar;
            sheet.Cells["A3"].Value = BaseCharPlusVs01;

            StyleReference(sheet.Cells["A2:A3"]);
            sheet.Cells.AutoFitColumns();

            for (int row = 1; row <= 3; row++)
            {
                sheet.Row(row).CustomHeight = true;
                sheet.Row(row).Height = 100;
            }
        }

        private static void StyleReference(ExcelRange range)
        {
            range.Style.Font.Name = ReproFontFamily;
            range.Style.Font.Size = ReproFontSize;
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
            range.Style.Indent = 0;
            range.Style.WrapText = false;
        }

        private static void StyleHeader(ExcelRange range)
        {
            range.Style.Font.Name = ReproFontFamily;
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