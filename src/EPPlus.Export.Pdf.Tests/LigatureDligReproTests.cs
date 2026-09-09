/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/07/2026         EPPlus Software AB           WP3/WP4 dlig ligature repro sheet
 *************************************************************************************************/
using EPPlus.Export.Pdf.Settings;
using EPPlus.Export.Pdf.Tests;
using EPPlus.Fonts.OpenType;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style;
using System;
using System.IO;

namespace EPPlusTest.PDF
{
    /// <summary>
    /// Visible before/after repro for the ligature feature-tag fix: EBGaramond's "dlig" feature
    /// merges "Th" into a single connected historical ligature glyph - a real, visually obvious
    /// effect, unlike "liga" which only covers the common ff/fi/fl set.
    ///
    /// Does not depend on a system font. EBGaramond-Regular.ttf is a repo test font (see
    /// EPPlus.Fonts.OpenType.Tests/Fonts) and is loaded via workbook.ConfigureFonts pointing at
    /// this test project's own Fonts folder, so both tests below are deterministic and CI-safe.
    ///
    /// The two tests are NOT symmetric in how they export, and that is unavoidable given the
    /// current public API:
    ///
    ///   - "Off" uses PdfTestBase.SaveAsPdf(sheet, name), since the default GsubFeatures
    ///     (Liga | Clig) is already what it needs.
    ///   - "On" needs GsubFeature.Dlig, which cannot be reached through the plain SaveAsPdf
    ///     overloads at all - ExcelWorksheet.SaveAsPdf always builds its own PdfPageSettings
    ///     internally (GetPdfSettings.GetPdfSettingsFromPrinterSettings), with no path from
    ///     ExcelPrinterSettings into GsubFeatures/GposFeatures. "On" therefore uses
    ///     PdfTestBase's PdfPageSettings-accepting SaveAsPdf overload instead.
    /// </summary>
    [TestClass]
    public class LigatureDligReproTests : PdfTestBase
    {
        private const string ReproFontFamily = "EB Garamond";
        private const float ReproFontSize = 48f;
        private const string ReproText = "Th";

        private static string FontsFolder => Path.Combine(AppContext.BaseDirectory, "Fonts");

        [TestMethod]
        public void Th_EBGaramond48_DligOff_DoesNotLigate()
        {
            using (var package = OpenPackage("WP3DligRepro_Off.xlsx", true))
            {
                var workbook = package.Workbook;

                workbook.ConfigureFonts(cfg =>
                {
                    cfg.FontDirectories.Add(FontsFolder);
                    cfg.SearchSystemDirectories = false;
                });

                var sheet = BuildReproSheet(workbook);

                // Default GsubFeatures is Liga | Clig - "dlig" is not requested, so this is the
                // "before" state: "Th" renders as two separate glyphs.
                SaveAsPdf(sheet, "WP3DligRepro_Off");

                SaveWorkbook("WP3DligRepro_Off.xlsx", package);
            }
        }

        [TestMethod]
        public void Th_EBGaramond48_DligOn_MergesIntoLigature()
        {
            using (var engine = new OpenTypeFontEngine(cfg =>
            {
                cfg.FontDirectories.Add(FontsFolder);
                cfg.SearchSystemDirectories = false;
            }))
            using (var package = OpenPackage("WP3DligRepro_On.xlsx", true))
            {
                var sheet = BuildReproSheet(package.Workbook);

                // GsubFeature.Dlig cannot be requested through the plain SaveAsPdf overloads -
                // see PdfTestBase's PdfPageSettings-accepting overload for why.
                var settings = new PdfPageSettings(engine)
                {
                    GsubFeatures = GsubFeature.Liga | GsubFeature.Clig | GsubFeature.Dlig
                };

                SaveAsPdf(sheet, "WP3DligRepro_On", settings);

                SaveWorkbook("WP3DligRepro_On.xlsx", package);
            }
        }

        /// <summary>
        /// Open WP3DligRepro_Off.pdf and WP3DligRepro_On.pdf side by side (both under
        /// _pdfPath, i.e. c:\epplusTest\Testoutput\PDF\). In "Off", "Th" renders as a plain T
        /// followed by a plain h. In "On" it renders as a single connected T-h glyph with a
        /// joining stroke - visibly narrower than the two separate letters. Before the fix, both
        /// files rendered identically, since "dlig" never reached the shaper regardless of what
        /// GsubFeatures was set to.
        /// </summary>
        private static ExcelWorksheet BuildReproSheet(ExcelWorkbook workbook)
        {
            var sheet = workbook.Worksheets.Add("Dlig");

            sheet.Cells["A1"].Value = ReproText;
            sheet.Cells["A1"].Style.Font.Name = ReproFontFamily;
            sheet.Cells["A1"].Style.Font.Size = ReproFontSize;
            sheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            sheet.Cells["A1"].Style.Indent = 0;

            sheet.Column(1).Width = 25;
            sheet.Row(1).CustomHeight = true;
            sheet.Row(1).Height = 70;

            return sheet;
        }
    }
}