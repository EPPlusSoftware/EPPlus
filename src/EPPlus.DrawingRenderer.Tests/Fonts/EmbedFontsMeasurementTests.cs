using System.Text.RegularExpressions;
using EPPlus.DrawingRenderer.Svg;
using EPPlus.Fonts.OpenType;
using EPPlusImageRenderer;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.DrawingRenderer.Tests.Fonts
{
    /// <summary>
    /// Step 4: <see cref="SvgRenderOptions.EmbedFonts"/> decides whether text is measured with the
    /// requested font or with its web substitute. <see cref="RenderContext.SubstituteWebFonts"/>
    /// carries the decision down to measurement.
    /// </summary>
    [TestClass]
    public class EmbedFontsMeasurementTests : FontTestBase
    {
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
                cfg.FontDirectories.Add(FontFolder);
                cfg.SearchSystemDirectories = false;
                configure?.Invoke(cfg);
            }));
            return p;
        }

        private static string FirstTspanFamily(string svg)
        {
            var m = Regex.Match(svg, "<tspan[^>]*?style=\"[^\"]*?font-family:((?:&[#\\w]+;|[^;\"])*)");
            Assert.IsTrue(m.Success, "No text run with a font-family was rendered.");
            return System.Net.WebUtility.HtmlDecode(m.Groups[1].Value).Split(',')[0].Trim();
        }

        private static string ShapeSvg(ExcelPackage p, string fontName, Action<SvgRenderOptions> options)
        {
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetSize(300, 200);
            shape.TextBody.Paragraphs.Add("Some text").TextRuns[0].SetFromFont(fontName, 11);
            return shape.ToSvg(options);
        }

        // -----------------------------------------------------------------------------------------
        // RenderContext
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        [DataRow(FontRenderTarget.Web, true, "Calibri")]
        [DataRow(FontRenderTarget.Web, false, "Aptos Narrow")]
        [DataRow(FontRenderTarget.Document, true, "Aptos Narrow")]
        [DataRow(FontRenderTarget.Document, false, "Aptos Narrow")]
        public void RenderContext_SubstitutesOnlyForWebWithFlag(FontRenderTarget target, bool substitute, string expected)
        {
            using (var engine = new OpenTypeFontEngine(cfg => cfg.SearchSystemDirectories = false))
            {
                var rc = new RenderContext(() => engine).ForTarget(target, substitute);

                Assert.AreEqual(target, rc.Target);
                Assert.AreEqual(substitute, rc.SubstituteWebFonts);
                Assert.AreEqual(expected, rc.GetFamilyForTarget("Aptos Narrow"));
            }
        }

        [TestMethod]
        public void RenderContext_ForTargetWithoutFlag_SubstitutesForWebOnly()
        {
            using (var engine = new OpenTypeFontEngine(cfg => cfg.SearchSystemDirectories = false))
            {
                var root = new RenderContext(() => engine);

                Assert.IsTrue(root.ForTarget(FontRenderTarget.Web).SubstituteWebFonts);
                Assert.IsFalse(root.ForTarget(FontRenderTarget.Document).SubstituteWebFonts);
                Assert.AreEqual("Calibri", root.ForTarget(FontRenderTarget.Web).GetFamilyForTarget("Aptos Narrow"));
            }
        }

        [TestMethod]
        public void RenderContext_Views_ShareTheRootEngineAndReturnRootWhenEqual()
        {
            using (var engine = new OpenTypeFontEngine(cfg => cfg.SearchSystemDirectories = false))
            {
                var root = new RenderContext(() => engine);
                var web = root.ForTarget(FontRenderTarget.Web, true);
                var webNoSubstitution = web.ForTarget(FontRenderTarget.Web, false);

                Assert.AreSame(root, web.ForTarget(FontRenderTarget.Document, false));
                Assert.AreSame(web, web.ForTarget(FontRenderTarget.Web, true));
                Assert.AreSame(root.FontEngine, webNoSubstitution.FontEngine);
            }
        }

        // -----------------------------------------------------------------------------------------
        // SvgRenderOptions.EmbedFonts
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void SvgRenderOptions_EmbedFonts_DefaultsToTrue()
        {
            Assert.IsTrue(new SvgRenderOptions().EmbedFonts);
        }

        [TestMethod]
        public void Shape_EmbedFonts_MeasuresAndWritesRequestedFont()
        {
            using (var p = CreatePackage())
            {
                var svg = ShapeSvg(p, "Aptos Narrow", o => o.EmbedFonts = true);
                Assert.AreEqual("'Aptos Narrow'", FirstTspanFamily(svg));
            }
        }

        [TestMethod]
        public void Shape_NoEmbedFonts_WebTarget_UsesSubstitute()
        {
            using (var p = CreatePackage())
            {
                var svg = ShapeSvg(p, "Aptos Narrow", o => o.EmbedFonts = false);
                Assert.AreEqual("'Calibri'", FirstTspanFamily(svg));
            }
        }

        [TestMethod]
        public void Shape_NoEmbedFonts_DocumentTarget_KeepsRequestedFont()
        {
            using (var p = CreatePackage())
            {
                var svg = ShapeSvg(p, "Aptos Narrow", o =>
                {
                    o.EmbedFonts = false;
                    o.FontTarget = FontRenderTarget.Document;
                });
                Assert.AreEqual("'Aptos Narrow'", FirstTspanFamily(svg));
            }
        }

        [TestMethod]
        public void Chart_EmbedFonts_MeasuresAndWritesRequestedFont()
        {
            using (var p = CreatePackage())
            {
                var chart = AddColumnChart(p, "Aptos Narrow");
                Assert.AreEqual("'Aptos Narrow'", FirstTspanFamily(chart.ToSvg(o => o.EmbedFonts = true)));
                Assert.AreEqual("'Calibri'", FirstTspanFamily(chart.ToSvg(o => o.EmbedFonts = false)));
            }
        }

        // -----------------------------------------------------------------------------------------
        // Axis width is measured with the same font as the axis labels
        // -----------------------------------------------------------------------------------------

        private static ExcelChart AddColumnChart(ExcelPackage p, string axisFont)
        {
            var ws = p.Workbook.Worksheets.Add("Sheet" + (p.Workbook.Worksheets.Count + 1));
            ws.Cells["A1:A4"].FillNumber(1, 1);
            ws.Cells["B1"].Value = 1000000;
            ws.Cells["B2"].Value = 2500000;
            ws.Cells["B3"].Value = 4000000;
            ws.Cells["B4"].Value = 5500000;
            var chart = ws.Drawings.AddChart("Chart1", eChartType.ColumnClustered);
            chart.Series.Add(ws.Cells["B1:B4"], ws.Cells["A1:A4"]);
            chart.SetSize(500, 300);
            chart.XAxis.Font.LatinFont = axisFont;
            chart.YAxis.Font.LatinFont = axisFont;
            return chart;
        }

        private static double VerticalAxisWidth(string substitute)
        {
            using (var p = CreatePackage(cfg =>
            {
                cfg.WebFontSubstitutions["Aptos Narrow"] = substitute;
                // Measure from serialized metrics, so the result does not depend on installed fonts.
                cfg.MetricsFallback = MetricsFallbackMode.Always;
            }))
            {
                var chart = AddColumnChart(p, "Aptos Narrow");
                var renderer = new ChartRenderer(chart, new SvgRenderOptions { EmbedFonts = false });
                Assert.IsNotNull(renderer.VerticalAxis, "Test setup: the chart has no vertical axis.");
                return renderer.VerticalAxis.Rectangle.Width;
            }
        }

        [TestMethod]
        public void Chart_VerticalAxisWidth_FollowsTheSubstitutedFont()
        {
            // Verdana is much wider than Calibri. Before the fix the axis width was measured with the
            // unsubstituted font, so it was the same in both cases.
            using (var engine = new OpenTypeFontEngine(cfg =>
            {
                cfg.SearchSystemDirectories = false;
                cfg.MetricsFallback = MetricsFallbackMode.Always;
            }))
            {
                if (engine.GetMeasurementShaper("Verdana").HasGlyphIds || engine.GetMeasurementShaper("Calibri").HasGlyphIds)
                    Assert.Inconclusive("Test setup: Verdana and Calibri are expected to have serialized metrics.");
            }

            var verdana = VerticalAxisWidth("Verdana");
            var calibri = VerticalAxisWidth("Calibri");

            Assert.IsTrue(verdana > calibri,
                $"Expected the axis to be wider with Verdana than with Calibri (Verdana {verdana}, Calibri {calibri}).");
        }
    }
}