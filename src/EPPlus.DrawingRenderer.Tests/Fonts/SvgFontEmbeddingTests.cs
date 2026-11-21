using System.Globalization;
using System.Text.RegularExpressions;
using EPPlus.DrawingRenderer.Svg;
using EPPlus.Fonts.OpenType;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.DrawingRenderer.Tests.Fonts
{
    /// <summary>
    /// Step 5: fonts embedded in the svg as subsets (@font-face), font-family stacks that start with
    /// the subsets, textLength on every text run, xml:space on text elements, the root text style,
    /// and text runs split where measurement used a glyph fallback that is written by name.
    /// The font engine only sees the test fonts, so the output does not depend on installed fonts.
    /// </summary>
    [TestClass]
    public class SvgFontEmbeddingTests : FontTestBase
    {
        private const string JapaneseText = "Sales あいう";

        [TestInitialize]
        public void Initialize()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
        }

        private OpenTypeFontEngine? _engine;

        [TestCleanup]
        public void Cleanup()
        {
            _engine?.Dispose();
            _engine = null;
        }

        private ExcelPackage CreatePackage(Action<IEpplusFontConfiguration>? configure = null)
        {
            _engine = new OpenTypeFontEngine(cfg =>
            {
                cfg.FontDirectories.Add(FontFolder);
                cfg.SearchSystemDirectories = false;
                configure?.Invoke(cfg);
            });
            var p = new ExcelPackage();
            p.Workbook.UseFontEngine(_engine);
            return p;
        }

        private static ExcelShape AddShape(ExcelPackage p, string text, string fontName, bool bold = false)
        {
            var ws = p.Workbook.Worksheets.Add("Sheet" + (p.Workbook.Worksheets.Count + 1));
            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetSize(300, 200);
            shape.TextBody.Paragraphs.Add(text).TextRuns[0].SetFromFont(fontName, 11, bold);
            return shape;
        }

        /// <summary>Hiragana is drawn with BIZ UDGothic, which an OnFontEmbedding callback may forbid.</summary>
        private static void JapaneseFallback(IEpplusFontConfiguration cfg, bool embedFallback)
        {
            RequireFont(TestFolderEngine, "BIZ UDGothic");
            cfg.SetScriptFallback(UnicodeScript.Hiragana, "BIZ UDGothic", "Yu Gothic", "Meiryo");
            if (!embedFallback)
            {
                cfg.OnFontEmbedding(info => info.FontName.StartsWith("BIZ UDGothic", StringComparison.OrdinalIgnoreCase)
                    ? FontEmbeddingDecision.Skip
                    : FontEmbeddingDecision.Default);
            }
        }

        private static List<string> FontFaceFamilies(string svg)
        {
            return Regex.Matches(svg, "@font-face\\{font-family:'([^']*)'")
                .Cast<Match>().Select(m => m.Groups[1].Value).ToList();
        }

        private static string[] Families(string fontFamilyAttribute)
        {
            return System.Net.WebUtility.HtmlDecode(fontFamilyAttribute)
                .Split(',').Select(f => f.Trim().Trim('\'')).ToArray();
        }

        /// <summary>The innermost tspans, the ones that hold text.</summary>
        private static MatchCollection TextTspans(string svg)
        {
            return Regex.Matches(svg, "<tspan([^>]*)>([^<]+)</tspan>");
        }

        private static string? Attribute(string attributes, string name)
        {
            var m = Regex.Match(attributes, "\\s" + Regex.Escape(name) + "=\"([^\"]*)\"");
            return m.Success ? m.Groups[1].Value : null;
        }

        /// <summary>The width in pixels, measured with the engine of the package the test created.</summary>
        private double MeasurePixels(string text, string family, FontSubFamily subFamily, float size)
        {
            var points = _engine!.GetMeasurementShaper(family, subFamily).ShapeLight(text).GetWidthInPoints(size);
            return points / 72d * 96d;
        }

        // -----------------------------------------------------------------------------------------
        // @font-face
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void EmbedFonts_WritesFontFaceWithDataUriInDefs()
        {
            RequireFont(TestFolderEngine, "Roboto");
            using (var p = CreatePackage())
            {
                var svg = AddShape(p, "Hello", "Roboto").ToSvg();

                var m = Regex.Match(svg,
                    "<defs><style>@font-face\\{font-family:'([A-Z]{6}\\+Roboto)';src:url\\(data:font/ttf;base64,([A-Za-z0-9+/=]+)\\) format\\('truetype'\\);font-weight:400;font-style:normal;\\}");
                Assert.IsTrue(m.Success, "No @font-face for Roboto first in <defs>.");

                var subset = TestFolderEngine.GetFromBytes(Convert.FromBase64String(m.Groups[2].Value));
                foreach (var c in "Hello")
                    Assert.IsTrue(subset.CmapTable.GetGlyphId(c) > 0, "The subset has no glyph for '" + c + "'.");
                Assert.AreEqual(0, subset.CmapTable.GetGlyphId('Z'), "The subset has a glyph that is not used.");
            }
        }

        [TestMethod]
        public void EmbedFonts_BoldRun_IsMeasuredAndEmbeddedBold()
        {
            RequireFont(TestFolderEngine, "Roboto", FontSubFamily.Bold);
            using (var p = CreatePackage())
            {
                var svg = AddShape(p, "Hello", "Roboto", bold: true).ToSvg();

                StringAssert.Contains(svg, "font-weight:700;font-style:normal;}");
                var run = TextTspans(svg)[0].Groups[1].Value;
                StringAssert.Contains(run, "font-weight=\"bold\"");
                Assert.AreEqual(MeasurePixels("Hello", "Roboto", FontSubFamily.Bold, 11),
                    double.Parse(Attribute(run, "textLength")!, CultureInfo.InvariantCulture), 0.01);
            }
        }

        [TestMethod]
        public void EmbedFonts_StackStartsWithTheSubset()
        {
            RequireFont(TestFolderEngine, "Roboto");
            using (var p = CreatePackage())
            {
                var svg = AddShape(p, "Hello", "Roboto").ToSvg();
                var subsetName = FontFaceFamilies(svg).Single();

                var families = Families(Attribute(TextTspans(svg)[0].Groups[1].Value, "font-family")!);
                Assert.AreEqual(subsetName, families[0]);
                Assert.AreEqual("Roboto", families[1]);
                Assert.AreEqual("sans-serif", families.Last());
            }
        }

        [TestMethod]
        public void NoEmbedFonts_WritesNoFontFace()
        {
            RequireFont(TestFolderEngine, "Roboto");
            using (var p = CreatePackage())
            {
                var svg = AddShape(p, "Hello", "Roboto").ToSvg(o => o.EmbedFonts = false);
                Assert.IsFalse(svg.Contains("@font-face"));
                Assert.IsFalse(svg.Contains("<style>"));
            }
        }

        [TestMethod]
        public void Chart_EmbedFonts_EveryStackHasADeclaredSubset()
        {
            RequireFont(TestFolderEngine, "Roboto");
            using (var p = CreatePackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1:A4"].FillNumber(1, 1);
                ws.Cells["B1:B4"].FillNumber(10, 10);
                var chart = ws.Drawings.AddChart("Chart1", eChartType.ColumnClustered);
                chart.Series.Add(ws.Cells["B1:B4"], ws.Cells["A1:A4"]);
                chart.Title.Text = "Title";
                chart.Title.Font.LatinFont = "Roboto";
                chart.XAxis.Font.LatinFont = "Roboto";
                chart.YAxis.Font.LatinFont = "Roboto";

                var svg = chart.ToSvg();
                var faces = FontFaceFamilies(svg);
                Assert.IsTrue(faces.Any(f => f.EndsWith("+Roboto")), "Roboto is not embedded.");

                var runs = TextTspans(svg);
                Assert.IsTrue(runs.Count > 0);
                // A font that is not installed (here the title's Aptos Narrow) is measured from metrics and
                // written by name, followed by an embedded safety net. Every stack therefore has a subset,
                // and every subset a stack names is declared.
                foreach (Match run in runs)
                {
                    var fontFamily = Attribute(run.Groups[1].Value, "font-family");
                    if (fontFamily == null)
                        continue;
                    var subsets = Families(fontFamily).Where(f => Regex.IsMatch(f, @"^[A-Z]{6}\+")).ToList();
                    Assert.IsTrue(subsets.Count > 0, "Text run '" + run.Groups[2].Value + "' has no subset in its stack.");
                    foreach (var subset in subsets)
                        CollectionAssert.Contains(faces, subset, "Undeclared subset in the stack of '" + run.Groups[2].Value + "'.");
                }
            }
        }

        // -----------------------------------------------------------------------------------------
        // textLength, xml:space, root
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void EveryTextRun_HasTextLengthEqualToTheMeasuredWidth()
        {
            RequireFont(TestFolderEngine, "Roboto");
            using (var p = CreatePackage())
            {
                var svg = AddShape(p, "This text is long enough to wrap over more than one line in the shape.", "Roboto").ToSvg();

                var runs = TextTspans(svg);
                Assert.IsTrue(runs.Count > 1, "Expected the text to wrap.");
                foreach (Match run in runs)
                {
                    var text = System.Net.WebUtility.HtmlDecode(run.Groups[2].Value);
                    var textLength = Attribute(run.Groups[1].Value, "textLength");
                    Assert.IsNotNull(textLength, "No textLength on '" + text + "'.");
                    Assert.AreEqual("spacing", Attribute(run.Groups[1].Value, "lengthAdjust"));
                    Assert.AreEqual(MeasurePixels(text, "Roboto", FontSubFamily.Regular, 11),
                        double.Parse(textLength!, CultureInfo.InvariantCulture), 0.01, "textLength of '" + text + "'.");
                }
            }
        }

        [TestMethod]
        public void TextElements_PreserveSpaceAndHaveNothingBetweenRuns()
        {
            RequireFont(TestFolderEngine, "Roboto");
            using (var p = CreatePackage())
            {
                var svg = AddShape(p, "This text is long enough to wrap over more than one line in the shape.", "Roboto").ToSvg();

                var texts = Regex.Matches(svg, "<text([^>]*)>(.*?)</text>", RegexOptions.Singleline);
                Assert.IsTrue(texts.Count > 0);
                foreach (Match text in texts)
                {
                    StringAssert.Contains(text.Groups[1].Value, "xml:space=\"preserve\"");
                    // Outside the tspans there may be no text, not even white space.
                    var outside = Regex.Replace(text.Groups[2].Value, "<tspan[^>]*>[^<]*</tspan>|<tspan[^>]*>|</tspan>", "");
                    Assert.AreEqual(string.Empty, outside, "Text outside the runs of a text element.");
                }
            }
        }

        [TestMethod]
        public void Root_HasViewBoxAndTextStyle()
        {
            RequireFont(TestFolderEngine, "Roboto");
            using (var p = CreatePackage())
            {
                var svg = AddShape(p, "Hello", "Roboto").ToSvg();
                var root = Regex.Match(svg, "<svg[^>]*>").Value;

                StringAssert.Contains(root, " viewBox=\"");
                Assert.IsFalse(root.Contains(" viewbox="));
                var style = Attribute(root, "style");
                Assert.IsNotNull(style, "The root has no style.");
                foreach (var rule in new[] { "font-kerning:normal", "font-variant-ligatures:common-ligatures contextual",
                    "letter-spacing:0", "word-spacing:0", "font-synthesis:none", "text-rendering:geometricPrecision" })
                {
                    StringAssert.Contains(style, rule);
                }
            }
        }

        // -----------------------------------------------------------------------------------------
        // Glyph fallback
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void EmbeddedGlyphFallback_IsOneRunWithBothSubsetsInTheStack()
        {
            RequireFont(TestFolderEngine, "Roboto");
            using (var p = CreatePackage(cfg => JapaneseFallback(cfg, embedFallback: true)))
            {
                var svg = AddShape(p, JapaneseText, "Roboto").ToSvg();

                var faces = FontFaceFamilies(svg);
                Assert.AreEqual(2, faces.Count, string.Join(", ", faces));
                var runs = TextTspans(svg);
                Assert.AreEqual(1, runs.Count, "An embedded fallback needs no split.");

                var families = Families(Attribute(runs[0].Groups[1].Value, "font-family")!);
                Assert.IsTrue(families[0].EndsWith("+Roboto"));
                Assert.IsTrue(families[1].EndsWith("+BIZ UDGothic"));
                Assert.AreEqual(MeasurePixels(JapaneseText, "Roboto", FontSubFamily.Regular, 11),
                    double.Parse(Attribute(runs[0].Groups[1].Value, "textLength")!, CultureInfo.InvariantCulture), 0.01);
            }
        }

        [TestMethod]
        public void GlyphFallbackByName_IsSplitIntoItsOwnTspan()
        {
            RequireFont(TestFolderEngine, "Roboto");
            using (var p = CreatePackage(cfg => JapaneseFallback(cfg, embedFallback: false)))
            {
                var svg = AddShape(p, JapaneseText, "Roboto").ToSvg();

                // Only Roboto is embedded; BIZ UDGothic is written by name.
                var faces = FontFaceFamilies(svg);
                Assert.AreEqual(1, faces.Count, string.Join(", ", faces));

                // The run is an outer tspan without textLength, holding one tspan per font.
                var outer = Regex.Match(svg, "<tspan([^>]*)>((?:<tspan[^>]*>[^<]*</tspan>)+)</tspan>");
                Assert.IsTrue(outer.Success, "The run was not split.");
                Assert.IsNull(Attribute(outer.Groups[1].Value, "textLength"));
                Assert.AreEqual(faces[0], Families(Attribute(outer.Groups[1].Value, "font-family")!)[0]);

                var segments = TextTspans(outer.Groups[2].Value);
                Assert.AreEqual(2, segments.Count);

                Assert.AreEqual("Sales ", segments[0].Groups[2].Value);
                Assert.IsNull(Attribute(segments[0].Groups[1].Value, "font-family"), "The Latin part inherits the run's stack.");

                Assert.AreEqual("あいう", segments[1].Groups[2].Value);
                var families = Families(Attribute(segments[1].Groups[1].Value, "font-family")!);
                CollectionAssert.AreEqual(new[] { "BIZ UDGothic", "Yu Gothic", "Meiryo" }, families.Take(3).ToArray(),
                    "The fallback, then the rest of the script's chain.");
                Assert.AreEqual("sans-serif", families.Last());

                foreach (Match segment in segments)
                {
                    Assert.AreEqual(MeasurePixels(segment.Groups[2].Value, "Roboto", FontSubFamily.Regular, 11),
                        double.Parse(Attribute(segment.Groups[1].Value, "textLength")!, CultureInfo.InvariantCulture), 0.01,
                        "textLength of '" + segment.Groups[2].Value + "'.");
                }
            }
        }

        [TestMethod]
        public void GlyphFallbackByName_RunStackListsTheScriptChain()
        {
            RequireFont(TestFolderEngine, "Roboto");
            using (var p = CreatePackage(cfg => JapaneseFallback(cfg, embedFallback: false)))
            {
                var svg = AddShape(p, JapaneseText, "Roboto").ToSvg();

                var outer = Regex.Match(svg, "<tspan([^>]*)><tspan");
                var families = Families(Attribute(outer.Groups[1].Value, "font-family")!).ToList();
                var fallback = families.IndexOf("BIZ UDGothic");
                Assert.IsTrue(fallback > 0, string.Join(", ", families));
                Assert.AreEqual("Yu Gothic", families[fallback + 1]);
                Assert.AreEqual("Meiryo", families[fallback + 2]);
            }
        }

        // -----------------------------------------------------------------------------------------
        // Without embedding (step 6)
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void NoEmbedFonts_GlyphFallback_IsSplitIntoItsOwnTspan()
        {
            RequireFont(TestFolderEngine, "Roboto");
            using (var p = CreatePackage(cfg => JapaneseFallback(cfg, embedFallback: true)))
            {
                var svg = AddShape(p, JapaneseText, "Roboto").ToSvg(o => o.EmbedFonts = false);
                Assert.IsFalse(svg.Contains("@font-face"));

                var outer = Regex.Match(svg, "<tspan([^>]*)>((?:<tspan[^>]*>[^<]*</tspan>)+)</tspan>");
                Assert.IsTrue(outer.Success, "The run was not split.");
                CollectionAssert.AreEqual(new[] { "Roboto", "BIZ UDGothic", "Yu Gothic", "Meiryo", "sans-serif" },
                    Families(Attribute(outer.Groups[1].Value, "font-family")!),
                    "The run's stack: the font, the glyph fallback and its script chain, a generic family.");

                var segments = TextTspans(outer.Groups[2].Value);
                Assert.AreEqual(2, segments.Count);
                Assert.IsNull(Attribute(segments[0].Groups[1].Value, "font-family"));
                CollectionAssert.AreEqual(new[] { "BIZ UDGothic", "Yu Gothic", "Meiryo", "Roboto", "sans-serif" },
                    Families(Attribute(segments[1].Groups[1].Value, "font-family")!));

                foreach (Match segment in segments)
                {
                    Assert.AreEqual(MeasurePixels(segment.Groups[2].Value, "Roboto", FontSubFamily.Regular, 11),
                        double.Parse(Attribute(segment.Groups[1].Value, "textLength")!, CultureInfo.InvariantCulture), 0.01,
                        "textLength of '" + segment.Groups[2].Value + "'.");
                }
            }
        }

        [TestMethod]
        public void NoEmbedFonts_FontNotInstalled_HasNoSafetyNet()
        {
            using (var p = CreatePackage())
            {
                // Calibri is not in the test folder. Without embedding there is no Archivo Narrow subset.
                var svg = AddShape(p, "Hello", "Calibri").ToSvg(o => { o.EmbedFonts = false; o.FontTarget = FontRenderTarget.Document; });

                Assert.IsFalse(svg.Contains("@font-face"));
                Assert.IsFalse(svg.Contains("Archivo Narrow"), "Archivo Narrow is only ever used embedded.");
                var families = Families(Attribute(TextTspans(svg)[0].Groups[1].Value, "font-family")!);
                Assert.AreEqual("Calibri", families[0]);
                Assert.AreEqual("sans-serif", families.Last());
            }
        }

        [TestMethod]
        public void NoEmbedFonts_DoesNotConsultTheEmbeddingCallback()
        {
            RequireFont(TestFolderEngine, "Roboto");
            int calls = 0;
            using (var p = CreatePackage(cfg => cfg.OnFontEmbedding(info => { calls++; return FontEmbeddingDecision.Default; })))
            {
                AddShape(p, "Hello", "Roboto").ToSvg(o => o.EmbedFonts = false);
                Assert.AreEqual(0, calls);

                AddShape(p, "Hello", "Roboto").ToSvg();
                Assert.IsTrue(calls > 0, "Test setup: the callback runs when fonts are embedded.");
            }
        }
    }
}