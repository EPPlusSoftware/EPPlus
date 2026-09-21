/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/09/2026         EPPlus Software AB           GSUB Multiple Substitution visual repro sheet
 *************************************************************************************************/
using EPPlus.Export.Pdf.Settings;
using EPPlus.Export.Pdf.Tests;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Tables;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Scripts;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using OfficeOpenXml;
using OfficeOpenXml.Export.PdfExport;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace EPPlus.Export.Pdf.Tests
{
    /// <summary>
    /// VISUAL repro for GSUB Lookup Type 2 (Multiple Substitution) actually being applied during
    /// shaping and preserved through subsetting/embedding - not an automated regression test,
    /// the point is to look at the page. The unit tests already cover the automated side.
    ///
    /// A synthetic Multiple Substitution lookup is appended to Roboto's own existing "liga"
    /// feature: the digit '5' is made to expand into the glyphs for 'H' and 'i'. That's a
    /// deliberately unmistakable effect - if it works, a cell containing the single character
    /// "5" should render as "Hi", identical in appearance to a cell where "Hi" was typed
    /// directly.
    ///
    /// What to look for in the exported PDF:
    ///
    ///   A2  "Hi"   Typed directly, two ordinary characters - the reference.
    ///   A3  "5"    A single digit in the underlying cell text. If Multiple Substitution works
    ///              end to end (shaping -> subsetting -> embedding), this renders identically to
    ///              A2. If it silently falls back, this renders as a plain "5".
    /// </summary>
    [TestClass]
    public class MultipleSubstitutionReproTests : PdfTestBase
    {
        private static string FontsFolder => Path.Combine(AppContext.BaseDirectory, "Fonts");

        private const string ReproFontFamily = "Roboto";
        private const float ReproFontSize = 48f;

        private static OpenTypeFontEngine CreateEngine()
        {
            return new OpenTypeFontEngine(cfg =>
            {
                cfg.FontDirectories.Add(FontsFolder);
                cfg.SearchSystemDirectories = false;
            });
        }

        /// <summary>
        /// Loads Roboto through the given engine's normal (cached) path and appends a synthetic
        /// Multiple Substitution lookup - '5' -> the glyphs for 'H' and 'i' - to its existing
        /// "liga" feature. Uses the cached load path deliberately, so the SAME mutated font
        /// instance is what the PDF export pipeline resolves later through this engine.
        /// </summary>
        private static void InjectDigitFiveExpandsToHi(OpenTypeFontEngine engine)
        {
            var font = engine.LoadFont(ReproFontFamily, FontSubFamily.Regular);

            if (!font.CmapTable.TryGetGlyphId('5', out ushort glyphIdFive))
                Assert.Fail("Test font is expected to contain '5'.");
            if (!font.CmapTable.TryGetGlyphId('H', out ushort glyphIdH))
                Assert.Fail("Test font is expected to contain 'H'.");
            if (!font.CmapTable.TryGetGlyphId('i', out ushort glyphIdI))
                Assert.Fail("Test font is expected to contain 'i'.");


            var multiSubst = new MultipleSubstSubTable
            {
                SubtableFormat = 1,
                Coverage = CoverageTableFormat2.CreateCoverageFormat2(new List<ushort> { glyphIdFive }),
                Sequences = new List<ushort[]> { new[] { glyphIdH, glyphIdI } }
            };

            var newLookup = new LookupTable
            {
                LookupType = 2,
                LookupFlag = 0,
                SubTables = new List<FontTableElement> { multiSubst }
            };

            int newLookupIndex = font.GsubTable.LookupList.Lookups.Count;
            font.GsubTable.LookupList.Lookups.Add(newLookup);

            AppendLookupToActiveLigaFeature(font, newLookupIndex);
        }

        /// <summary>
        /// Returns the index of a "liga" FeatureRecord that is actually REACHABLE from the
        /// "latn" script, and appends <paramref name="lookupIndex"/> to it.
        ///
        /// This matters: Roboto defines "liga" three times (once per script grouping), and only
        /// ONE of those FeatureRecords is listed in the latn LangSys. Picking the first "liga"
        /// by tag alone lands on a record the script-aware feature resolver correctly filters
        /// out, so the injected lookup would never run - the test would fail for a reason that
        /// has nothing to do with the code under test.
        /// </summary>
        private static void AppendLookupToActiveLigaFeature(OpenTypeFont font, int lookupIndex)
        {
            var activeIndices = ScriptFeatureResolver.GetActiveFeatureIndices(
                font.GsubTable.ScriptList, "latn", null);
            Assert.IsNotNull(activeIndices, "Test font is expected to have a ScriptList.");

            var records = font.GsubTable.FeatureList.FeatureRecords;
            int ligaIndex = -1;
            for (int i = 0; i < records.Count; i++)
            {
                if (records[i].FeatureTag.Value == "liga" && activeIndices.Contains(i))
                {
                    ligaIndex = i;
                    break;
                }
            }

            Assert.IsTrue(ligaIndex >= 0, "Test font is expected to define a 'liga' feature reachable from the latn script.");

            var featureTable = records[ligaIndex].FeatureTable;
            var oldIndices = featureTable.LookupListIndices ?? new ushort[0];
            var newIndices = new ushort[oldIndices.Length + 1];
            oldIndices.CopyTo(newIndices, 0);
            newIndices[oldIndices.Length] = (ushort)lookupIndex;
            featureTable.LookupListIndices = newIndices;
            featureTable.LookupCount = (ushort)newIndices.Length;
        }


        [TestMethod]
        public void MultipleSubstitution_Roboto_ReproSheet()
        {
            using (var package = OpenPackage("MultipleSubstitutionRepro.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("MultipleSubstitution");
                BuildReproSheet(sheet);

                var engine = CreateEngine();
                InjectDigitFiveExpandsToHi(engine);

                var settings = new PdfPageSettings(engine);

                // Written to disk for visual inspection - this is the whole point of the test.
                SaveAsPdf(sheet, "MultipleSubstitutionRepro", settings);

                // Also export to a stream so the test fails loudly if the export itself breaks,
                // rather than silently writing an unreadable file.
                using (var stream = new MemoryStream())
                {
                    new PdfCatalog(settings, sheet).Save(stream);
                    AssertLooksLikePdf(stream.ToArray());
                }

                SaveWorkbook("MultipleSubstitutionRepro.xlsx", package);
            }
        }

        private static void BuildReproSheet(ExcelWorksheet sheet)
        {
            sheet.Cells["A1"].Value = "Compare A2 (typed \"Hi\") vs A3 (typed \"5\") - both should look the same";
            StyleHeader(sheet.Cells["A1"]);

            sheet.Cells["A2"].Value = "Hi";
            sheet.Cells["A3"].Value = "5";

            StyleReference(sheet.Cells["A2:A3"]);
            sheet.Cells.AutoFitColumns();

            for (int row = 1; row <= 3; row++)
            {
                sheet.Row(row).CustomHeight = true;
                sheet.Row(row).Height = 70;
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