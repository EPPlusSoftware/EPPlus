using System.Text.RegularExpressions;
using EPPlus.DrawingRenderer.Svg;
using EPPlus.Fonts.OpenType;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Export.ImageRenderer.Tests.Fonts
{
    /// <summary>
    /// Verifies that web font substitution reaches the SVG output. The font engine does not search
    /// system directories, so the output is independent of the fonts installed on the machine.
    /// </summary>
    [TestClass]
    public class WebFontSubstitutionSvgTests : TestBase
    {
        private const string SampleText = "This is a line chart with data from the table below. The chart is exported as SVG when exporting to HTML.";

        [TestInitialize]
        public void Initialize()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
        }

        private static ExcelPackage CreatePackage(Action<IEpplusFontConfiguration>? configure = null)
        {
            var p = new ExcelPackage();
            p.Workbook.UseFontEngine(new OpenTypeFontEngine(cfg =>
            {
                cfg.SearchSystemDirectories = false;
                if (configure != null)
                {
                    configure(cfg);
                }
            }));
            return p;
        }

        private static ExcelShape AddShapeWithText(ExcelPackage p, string fontName)
        {
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetSize(300, 200);
            var paragraph = shape.TextBody.Paragraphs.Add(SampleText);
            paragraph.TextRuns[0].SetFromFont(fontName, 11);
            return shape;
        }

        /// <summary>
        /// Returns the first family of the font-family attribute on every text run (tspan) in the svg.
        /// The paragraph level (text element) carries the paragraph default font, which runs may override.
        /// </summary>
        private static List<string> GetRunFontFamilies(string svg)
        {
            var result = new List<string>();
            foreach (Match m in Regex.Matches(svg, "<tspan[^>]*?font-family=\"([^\"]*)\""))
            {
                var first = m.Groups[1].Value.Split(',')[0].Trim().Trim('\'', '"');
                result.Add(first);
            }
            return result;
        }

        private static void AssertAllFamilies(string svg, string expectedFamily)
        {
            var matches = Regex.Matches(svg, "<tspan[^>]*?font-family=\"([^\"]*)\"[^>]*>([^<]*)</tspan>");
            Assert.IsTrue(matches.Count > 0, "No text run with a font-family was rendered.");
            foreach (Match m in matches)
            {
                var family = m.Groups[1].Value.Split(',')[0].Trim().Trim('\'', '"');
                Assert.AreEqual(expectedFamily, family, "Text run '" + m.Groups[2].Value + "'");
            }
        }

        // -----------------------------------------------------------------------------------------
        // Shapes
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void Shape_OfficeCloudFont_IsSubstituted()
        {
            using (var p = CreatePackage())
            {
                var svg = AddShapeWithText(p, "Aptos Narrow").ToSvg();

                AssertAllFamilies(svg, "Calibri");
                Assert.IsFalse(svg.Contains("Aptos Narrow"), "The original font name leaked into the svg.");
            }
        }

        [TestMethod]
        public void Shape_FontNotInTable_IsUnchanged()
        {
            using (var p = CreatePackage())
            {
                var svg = AddShapeWithText(p, "Arial").ToSvg();

                AssertAllFamilies(svg, "Arial");
            }
        }

        [TestMethod]
        public void Shape_DocumentTarget_KeepsOriginalFont()
        {
            using (var p = CreatePackage())
            {
                var svg = AddShapeWithText(p, "Aptos Narrow").ToSvg(o => o.FontTarget = FontRenderTarget.Document);

                AssertAllFamilies(svg, "Aptos Narrow");
            }
        }

        [TestMethod]
        public void Shape_UserSubstitution_IsUsed()
        {
            using (var p = CreatePackage(cfg => cfg.WebFontSubstitutions["Aptos Narrow"] = "Verdana"))
            {
                var svg = AddShapeWithText(p, "Aptos Narrow").ToSvg();

                AssertAllFamilies(svg, "Verdana");
            }
        }

        // -----------------------------------------------------------------------------------------
        // Charts
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void Chart_OfficeCloudFont_IsSubstitutedInAllText()
        {
            using (var p = CreatePackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                LoadItemData(ws);
                var chart = ws.Drawings.AddChart("Chart1", eChartType.Line);
                chart.Series.Add(ws.Cells["N2:N11"], ws.Cells["K2:K11"]);
                chart.SetSize(600, 400);

                chart.Title.Text = "Chart title";
                chart.Title.Font.LatinFont = "Aptos Narrow";
                chart.XAxis.Font.LatinFont = "Aptos Narrow";
                chart.YAxis.Font.LatinFont = "Aptos Narrow";
                chart.Legend.Font.LatinFont = "Aptos Narrow";

                var svg = chart.ToSvg();

                //Covers title, axis labels and legend. A failure here points to a text path
                //that bypasses the render context's target.
                AssertAllFamilies(svg, "Calibri");
                Assert.IsFalse(svg.Contains("Aptos Narrow"), "The original font name leaked into the svg.");
            }
        }

        /// <summary>
        /// Measurement is independent of installed fonts: no system fonts and metrics always preferred.
        /// </summary>
        private static void Deterministic(IEpplusFontConfiguration cfg)
        {
            cfg.MetricsFallback = MetricsFallbackMode.Always;
        }

        /// <summary>
        /// Returns the text of every text run (tspan) in document order. For wrapped text each run is a line.
        /// </summary>
        private static List<string> GetRunTexts(string svg)
        {
            var result = new List<string>();
            foreach (Match m in Regex.Matches(svg, "<tspan[^>]*>([^<]*)</tspan>"))
            {
                result.Add(m.Groups[1].Value);
            }
            return result;
        }

        private static ExcelChart AddChartWithFont(ExcelPackage p, string fontName)
        {
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            LoadItemData(ws);
            var chart = ws.Drawings.AddChart("Chart1", eChartType.Line);
            chart.Series.Add(ws.Cells["N2:N11"], ws.Cells["K2:K11"]);
            chart.SetSize(400, 300);

            chart.Title.Text = "A long chart title that should wrap into more than one line";
            chart.Title.Font.LatinFont = fontName;
            chart.XAxis.Font.LatinFont = fontName;
            chart.YAxis.Font.LatinFont = fontName;
            chart.Legend.Font.LatinFont = fontName;
            return chart;
        }

        // -----------------------------------------------------------------------------------------
        // Measurement follows the substitution
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void Shape_WebSubstitution_WrapsAsSubstituteFont()
        {
            string webSvg, documentSvg;
            using (var p = CreatePackage(Deterministic))
            {
                webSvg = AddShapeWithText(p, "Aptos Narrow").ToSvg();
            }
            using (var p = CreatePackage(Deterministic))
            {
                documentSvg = AddShapeWithText(p, "Calibri").ToSvg(o => o.FontTarget = FontRenderTarget.Document);
            }

            var webLines = GetRunTexts(webSvg);
            Assert.IsTrue(webLines.Count > 1, "The text must wrap for the test to verify line breaking.");
            CollectionAssert.AreEqual(GetRunTexts(documentSvg), webLines);
        }

        [TestMethod]
        public void Shape_WrapComparison_DetectsDifferentFont()
        {
            //Control for the test above: a clearly wider font must break differently,
            //otherwise the comparison would pass regardless of which font is measured.
            string webSvg, documentSvg;
            using (var p = CreatePackage(Deterministic))
            {
                webSvg = AddShapeWithText(p, "Aptos Narrow").ToSvg();
            }
            using (var p = CreatePackage(Deterministic))
            {
                documentSvg = AddShapeWithText(p, "Verdana").ToSvg(o => o.FontTarget = FontRenderTarget.Document);
            }

            CollectionAssert.AreNotEqual(GetRunTexts(documentSvg), GetRunTexts(webSvg));
        }

        [TestMethod]
        public void Chart_WebSubstitution_LaysOutAsSubstituteFont()
        {
            string webSvg, documentSvg;
            using (var p = CreatePackage(Deterministic))
            {
                webSvg = AddChartWithFont(p, "Aptos Narrow").ToSvg();
            }
            using (var p = CreatePackage(Deterministic))
            {
                documentSvg = AddChartWithFont(p, "Calibri").ToSvg(o => o.FontTarget = FontRenderTarget.Document);
            }

            //Covers title wrapping, axis label selection and legend. Any text path that measures
            //with the original font instead of the substitute makes the sequences differ.
            CollectionAssert.AreEqual(GetRunTexts(documentSvg), GetRunTexts(webSvg));
        }
    }
}