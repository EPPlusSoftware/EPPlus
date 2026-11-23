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
    /// The font-family written to svg is a CSS stack: the measured family, its fallback chain
    /// (user-configured, then built-in) and a generic family from the measured font.
    ///
    /// The engines here search only the test font folder, so the expected stacks do not depend
    /// on the fonts installed on the machine.
    /// </summary>
    [TestClass]
    public class SvgFontFamilyStackTests : FontTestBase
    {
        private const string CalibriChain = "'Carlito', 'Liberation Sans', 'Arial', 'Helvetica'";

        [TestInitialize]
        public void Initialize()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
        }

        private Action<SvgRenderOptions> GetRenderOptions(FontRenderTarget target = FontRenderTarget.Document)
        {
            return o =>
            {
                o.FontTarget = target;
                o.EmbedFonts = false;
            };
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

        private static ExcelShape AddShape(ExcelPackage p, string fontName)
        {
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetSize(300, 200);
            var paragraph = shape.TextBody.Paragraphs.Add("Text in " + fontName);
            paragraph.TextRuns[0].SetFromFont(fontName, 11);
            return shape;
        }

        /// <summary>
        /// The font-family attribute of every tspan, decoded from xml.
        /// </summary>
        private static List<string> GetTspanFontFamilies(string svg)
        {
            return Regex.Matches(svg, "<tspan[^>]*?style=\"[^\"]*?font-family:((?:&[#\\w]+;|[^;\"])*)")
                .Cast<Match>()
                .Select(m => System.Net.WebUtility.HtmlDecode(m.Groups[1].Value))
                .ToList();
        }

        private static string GetSingleTspanFontFamily(string svg)
        {
            var families = GetTspanFontFamilies(svg);
            Assert.AreEqual(1, families.Count, "Expected one text run.");
            return families[0];
        }

        // -----------------------------------------------------------------------------------------
        // Shapes
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void Shape_FontWithoutChain_IsFamilyAndGeneric()
        {
            using (var p = CreatePackage())
            {
                var svg = AddShape(p, "Roboto").ToSvg(GetRenderOptions());
                Assert.AreEqual("'Roboto', sans-serif", GetSingleTspanFontFamily(svg));
            }
        }

        [TestMethod]
        public void Shape_FontWithBuiltinChain_ListsChain()
        {
            using (var p = CreatePackage())
            {
                // Calibri is not in the test folder; measurement falls through to the embedded
                // last-resort font, which is sans-serif.
                var svg = AddShape(p, "Calibri").ToSvg(GetRenderOptions());
                Assert.AreEqual("'Calibri', " + CalibriChain + ", sans-serif", GetSingleTspanFontFamily(svg));
            }
        }

        [TestMethod]
        public void Shape_UserChain_ComesFirst()
        {
            using (var p = CreatePackage(cfg => cfg.FontFallbacks["Calibri"] = new[] { "Open Sans" }))
            {
                var svg = AddShape(p, "Calibri").ToSvg(GetRenderOptions());
                Assert.AreEqual("'Calibri', 'Open Sans', " + CalibriChain + ", sans-serif", GetSingleTspanFontFamily(svg));
            }
        }

        [TestMethod]
        public void Shape_MonospaceFont_EndsWithMonospace()
        {
            using (var p = CreatePackage())
            {
                RequireFont(TestFolderEngine, "BIZ UDGothic");
                var svg = AddShape(p, "BIZ UDGothic").ToSvg(GetRenderOptions());
                Assert.AreEqual("'BIZ UDGothic', monospace", GetSingleTspanFontFamily(svg));
            }
        }

        [TestMethod]
        public void Shape_WebTarget_StackStartsWithSubstitute()
        {
            using (var p = CreatePackage())
            {
                // Aptos Narrow is substituted with Calibri for the web before measurement, so the
                // stack is Calibri's.
                var svg = AddShape(p, "Aptos Narrow").ToSvg(GetRenderOptions(FontRenderTarget.Web));
                Assert.AreEqual("'Calibri', " + CalibriChain + ", sans-serif", GetSingleTspanFontFamily(svg));
            }
        }

        [TestMethod]
        public void Shape_DocumentTarget_StackStartsWithRequestedFont()
        {
            using (var p = CreatePackage())
            {
                var svg = AddShape(p, "Aptos Narrow").ToSvg(GetRenderOptions());
                Assert.AreEqual(
                    "'Aptos Narrow', 'Liberation Sans Narrow', 'Arial Narrow', 'Carlito', 'Liberation Sans', 'Arial', sans-serif",
                    GetSingleTspanFontFamily(svg));
            }
        }

        [TestMethod]
        public void Shape_NameWithQuote_IsEscaped()
        {
            using (var p = CreatePackage())
            {
                var svg = AddShape(p, "Bob's \"Font\"").ToSvg(GetRenderOptions());

                // xml: the double quote is an entity so the attribute stays well-formed.
                StringAssert.Contains(svg, "font-family:'Bob\\'s &quot;Font&quot;', sans-serif;");
                // css after xml decoding: the single quote is escaped inside the quoted name.
                Assert.AreEqual("'Bob\\'s \"Font\"', sans-serif", GetSingleTspanFontFamily(svg));
            }
        }

        [TestMethod]
        public void Shape_NoMsFontService()
        {
            using (var p = CreatePackage())
            {
                var svg = AddShape(p, "Calibri").ToSvg(GetRenderOptions(FontRenderTarget.Web));
                Assert.IsFalse(svg.Contains("_MSFontService"), "The svg still contains _MSFontService.");
            }
        }

        [TestMethod]
        public void Shape_ParagraphAndGroupFonts_AreStacks()
        {
            using (var p = CreatePackage())
            {
                var svg = AddShape(p, "Calibri").ToSvg(GetRenderOptions());

                // Every font-family in the svg, on text, tspan and g, ends with a generic family. The group writes
                // it as an attribute, text and tspan in the style attribute. A font-family in a style attribute is
                // only matched inside style="...", so the @font-face rules in the style element are not counted;
                // xml entities such as &quot; are skipped, since they contain a semicolon.
                var all = Regex.Matches(svg, "(?:\\bfont-family=\"([^\"]*)\"|\\bstyle=\"[^\"]*?font-family:((?:&[#\\w]+;|[^;\"])*))")
                    .Cast<Match>()
                    .Select(m => m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value)
                    .ToList();
                Assert.IsTrue(all.Count >= 2, "Expected font-family on more than the text runs.");
                foreach (var family in all)
                {
                    Assert.IsTrue(family.EndsWith(", sans-serif") || family.EndsWith(", serif") || family.EndsWith(", monospace"),
                        "Not a stack: " + family);
                }
            }
        }

        // -----------------------------------------------------------------------------------------
        // Charts
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void Chart_AllTextRuns_AreStacks()
        {
            using (var p = CreatePackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1:A4"].FillNumber(1, 1);
                ws.Cells["B1:B4"].FillNumber(10, 5);
                var chart = ws.Drawings.AddChart("Chart1", eChartType.ColumnClustered);
                chart.Series.Add(ws.Cells["B1:B4"], ws.Cells["A1:A4"]);
                chart.SetSize(400, 300);
                chart.Title.Text = "Chart title";
                chart.Title.Font.LatinFont = "Calibri";
                chart.XAxis.Font.LatinFont = "Calibri";
                chart.YAxis.Font.LatinFont = "Calibri";

                var svg = chart.ToSvg(GetRenderOptions());

                Assert.IsFalse(svg.Contains("_MSFontService"), "The svg still contains _MSFontService.");
                var families = GetTspanFontFamilies(svg);
                Assert.IsTrue(families.Count > 0, "No text run was rendered.");
                foreach (var family in families.Where(f => f.StartsWith("'Calibri'")))
                {
                    Assert.AreEqual("'Calibri', " + CalibriChain + ", sans-serif", family);
                }
                Assert.IsTrue(families.Any(f => f.StartsWith("'Calibri'")), "No text run in Calibri.");
            }
        }
    }
}
