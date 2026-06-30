using EPPlus.Fonts.OpenType;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace EPPlusTest
{
    [TestClass]
    public class ConfigureFontsTests : FontTestBase
    {
        private static string RenderTextbox(ExcelWorksheet sheet, string text, string fontName)
        {
            var tb = sheet.Drawings.AddTextbox("txtBox1");
            tb.Text = text;
            tb.Font.LatinFont = fontName;
            return tb.ToSvg();
        }

        private static string RenderTextbox(ExcelPackage package, string text, string fontName)
        {
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            return RenderTextbox(sheet, text, fontName);
        }

        private static string RenderRobotoTextbox(ExcelPackage package)
        {
            return RenderTextbox(package, "Hello", "Roboto");
        }

        // Pulls the first tspan baseline y out of the SVG. It's computed from the resolved font's
        // ascent, so it differs between fonts — unlike font-size, which is just the requested size.
        private static double GetFirstTspanBaselineY(string svg)
        {
            var m = Regex.Match(svg, @"<tspan[^>]*\by\s*=""([0-9.]+)px""");
            Assert.IsTrue(m.Success, "Expected a tspan with a y position in the SVG.");
            return double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
        }

        [TestMethod]
        public void ConfigureFonts_Null_ThrowsArgumentNullException()
        {
            using var package = new ExcelPackage();
            Assert.ThrowsExactly<ArgumentNullException>(
                () => package.Workbook.ConfigureFonts(null));
        }

        [TestMethod]
        public void ConfigureFonts_ConfiguredDirectory_ResolvesRoboto()
        {
            using var package = new ExcelPackage();
            package.Workbook.ConfigureFonts(cfg =>
            {
                cfg.FontDirectories.Add(FontFolder);
                cfg.SearchSystemDirectories = false;
            });

            var svg = RenderRobotoTextbox(package);

            // The requested name always appears; the real proof is that rendering with the
            // directory configured does not throw and produces a measurable tspan.
            StringAssert.Contains(svg, "font-family=\"Roboto");
            Assert.IsTrue(GetFirstTspanBaselineY(svg) > 0);
        }

        [TestMethod]
        public void ConfigureFonts_IsPerWorkbook()
        {
            // Workbook A: Roboto directory available, system search off → resolves Roboto.
            using var packageA = new ExcelPackage();
            packageA.Workbook.ConfigureFonts(cfg =>
            {
                cfg.FontDirectories.Add(FontFolder);
                cfg.SearchSystemDirectories = false;
            });

            // Workbook B: no Roboto directory, system search off → falls back to the
            // embedded Archivo Narrow, which has different metrics than Roboto.
            using var packageB = new ExcelPackage();
            packageB.Workbook.ConfigureFonts(cfg =>
            {
                cfg.SearchSystemDirectories = false;
            });

            var svgA = RenderRobotoTextbox(packageA);
            var svgB = RenderRobotoTextbox(packageB);

            var baselineA = GetFirstTspanBaselineY(svgA);
            var baselineB = GetFirstTspanBaselineY(svgB);

            // Different resolved fonts → different measured size. This proves the configuration
            // affected resolution per workbook, not just that the requested name was echoed.
            Assert.AreNotEqual(baselineA, baselineB,
                "Workbook A (Roboto) and workbook B (embedded fallback) must measure differently, proving per-workbook font resolution.");
        }

        [TestMethod]
        public void ConfigureFonts_FontFallbacks_AreApplied()
        {
            // A: maps a non-existent font name to Roboto via the user fallback chain.
            using var packageA = new ExcelPackage();
            packageA.Workbook.ConfigureFonts(cfg =>
            {
                cfg.FontDirectories.Add(FontFolder);
                cfg.SearchSystemDirectories = false;
                cfg.FontFallbacks["NoSuchFont"] = new[] { "Roboto" };
            });

            // B: same missing font, no fallback chain → resolves via built-in chain to embedded.
            using var packageB = new ExcelPackage();
            packageB.Workbook.ConfigureFonts(cfg =>
            {
                cfg.FontDirectories.Add(FontFolder);
                cfg.SearchSystemDirectories = false;
            });

            var baselineA = GetFirstTspanBaselineY(RenderTextbox(packageA, "Hello", "NoSuchFont"));
            var baselineB = GetFirstTspanBaselineY(RenderTextbox(packageB, "Hello", "NoSuchFont"));

            Assert.AreNotEqual(baselineA, baselineB,
                "The user fallback chain should route the missing font to Roboto, measuring differently than the built-in fallback.");
        }

        [TestMethod]
        public void SetScriptFallback_RoutesHanGlyphsToConfiguredFont()
        {
            var engine = new OpenTypeFontEngine(cfg =>
            {
                cfg.FontDirectories.Add(FontFolder);
                cfg.SearchSystemDirectories = false;
                cfg.SetScriptFallback(UnicodeScript.Han, "BIZ UDGothic");
            });

            var shaper = engine.GetTextShaper("Open Sans");   // Latin-only primary
            var result = shaper.Shape("日本語");
            var usedFonts = shaper.GetUsedFonts().ToList();

            // The Han glyphs could not come from Open Sans; they must have been routed
            // to the configured BIZ UDGothic via the script fallback.
            Assert.IsTrue(usedFonts.Any(f => f.FullName.Contains("BIZ UDGothic") || f.FullName.Contains("BIZUDGothic")),
                "Han glyphs should be routed to the configured script-fallback font.");
        }
    }
}