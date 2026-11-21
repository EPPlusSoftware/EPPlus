using System.Drawing;
using System.Globalization;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using EPPlus.Fonts.OpenType;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.DrawingRenderer.Tests.Text
{
    /// <summary>
    /// Manual verification of text wrapping against Excel, before changing the wrap code.
    ///
    /// Writes WrapExcelComparison.xlsx with one shape per case, an svg per shape and an html page
    /// with all svgs, plus a text dump of every tspan (x and text). Open the workbook in Excel and
    /// compare each shape with its svg.
    ///
    /// The font is Calibri, read from the system font directories, so the test is Inconclusive
    /// where Calibri is not installed. The svgs are exported for the document target so Calibri is
    /// not substituted.
    ///
    /// Expected, from reading the wrap code:
    ///   A, B, E - line breaks and positions match Excel; the trailing space on wrapped lines
    ///             only affects run.Width internally, never a visible position.
    ///   C, D    - when a word is broken inside (no space to break at), the first run on the new
    ///             line is measured one character too narrow (the character moved to the new
    ///             line is not counted), so the next run on that line starts too far left and
    ///             overlaps it.
    /// </summary>
    [TestClass]
    public class WrapExcelComparisonTests : TestBase
    {
        private const string FontName = "Calibri";
        private const float FontSize = 11f;
        private const string OutputFolder = "svg\\WrapExcelComparison\\";

        private static readonly Color[] RunColors =
        {
            Color.Black, Color.Red, Color.Blue, Color.DarkGreen, Color.DarkOrange, Color.Purple
        };

        [TestMethod]
        public void WrapComparison_Calibri()
        {
            var engine = new OpenTypeFontEngine(cfg => cfg.SearchSystemDirectories = true);
            if (engine.GetFontAvailability(FontName) != FontAvailability.Exact
                || engine.GetFontAvailability(FontName, FontSubFamily.Bold) != FontAvailability.Exact)
            {
                Assert.Inconclusive(FontName + " (Regular and Bold) must be installed for this comparison.");
            }

            var cases = new List<ExcelShape>();
            var dump = new StringBuilder();

            using (var p = OpenPackage("WrapExcelComparison.xlsx", true))
            {
                p.Workbook.UseFontEngine(engine);
                var ws = p.Workbook.Worksheets.Add("Wrap");

                // A: right aligned, several lines broken at spaces.
                cases.Add(AddCase(ws, "A_RightAligned", 0, 150, eTextAlignment.Right,
                    R("The quick brown fox jumps over the lazy dog and keeps on running")));

                // B: the same text centered.
                cases.Add(AddCase(ws, "B_Centered", 1, 150, eTextAlignment.Center,
                    R("The quick brown fox jumps over the lazy dog and keeps on running")));

                // C: one long word broken inside, followed directly (no space) by a second run in
                // another colour. The second run lands on a line that starts with the word's tail.
                cases.Add(AddCase(ws, "C_MidWordBreak_TwoColours", 2, 60, eTextAlignment.Left,
                    R("Supercalifragilistic"), R("XYZ"), R("abc")));

                // D: as C, but the second run is bold and larger.
                cases.Add(AddCase(ws, "D_MidWordBreak_TwoFonts", 3, 60, eTextAlignment.Left,
                    R("Supercalifragilistic"), R("XYZ", bold: true, size: 14f), R("abc")));

                // E: break at a space on a fragment boundary, with more runs on the next line.
                cases.Add(AddCase(ws, "E_FragmentBoundary", 4, 70, eTextAlignment.Left,
                    R("Hello "), R("World "), R("again "), R("and again")));

                var html = new StringBuilder();
                html.AppendLine("<!doctype html><html><head><meta charset=\"utf-8\"><title>Wrap comparison</title>");
                html.AppendLine("<style>body{font-family:sans-serif} .case{display:inline-block;vertical-align:top;margin:12px} img{border:1px dashed #999}</style></head><body>");

                foreach (var shape in cases)
                {
                    var svg = shape.ToSvg(o => o.FontTarget = FontRenderTarget.Document);
                    var svgFile = GetOutputFile(OutputFolder, shape.Name + ".svg");
                    File.WriteAllText(svgFile.FullName, svg);

                    html.AppendLine($"<div class=\"case\"><div>{WebUtility.HtmlEncode(shape.Name)}</div><img src=\"{shape.Name}.svg\"></div>");

                    dump.AppendLine(shape.Name);
                    foreach (var tspan in GetTspans(engine, svg))
                    {
                        dump.AppendLine("  " + tspan);
                    }
                    dump.AppendLine();
                }

                html.AppendLine("</body></html>");
                File.WriteAllText(GetOutputFile(OutputFolder, "index.html").FullName, html.ToString());
                File.WriteAllText(GetOutputFile(OutputFolder, "tspans.txt").FullName, dump.ToString());

                SaveAndCleanup(p);
            }

            Console.WriteLine(dump.ToString());
        }

        private sealed class RunSpec
        {
            public string Text = "";
            public bool Bold;
            public float Size = FontSize;
        }

        private static RunSpec R(string text, bool bold = false, float size = FontSize)
        {
            return new RunSpec { Text = text, Bold = bold, Size = size };
        }

        /// <summary>
        /// Adds a rectangle with black text on white, no insets, at row 2 + index * 8.
        /// Each run gets its own colour so run boundaries are visible.
        /// </summary>
        private static ExcelShape AddCase(ExcelWorksheet ws, string name, int index, int widthPx, eTextAlignment alignment, params RunSpec[] runs)
        {
            var shape = ws.Drawings.AddShape(name, eShapeStyle.Rect);
            shape.SetPosition(1 + index * 8, 0, 1, 0);
            shape.SetSize(widthPx, 120);

            shape.Fill.Style = eFillStyle.SolidFill;
            shape.Fill.Color = Color.White;
            shape.Border.Fill.Style = eFillStyle.SolidFill;
            shape.Border.Fill.Color = Color.Gray;

            shape.TextBody.LeftInsert = 0;
            shape.TextBody.RightInsert = 0;
            shape.TextBody.TopInsert = 0;
            shape.TextBody.BottomInsert = 0;
            shape.TextBody.Anchor = eTextAnchoringType.Top;

            for (int i = 0; i < runs.Length; i++)
            {
                var rt = shape.RichText.Add(runs[i].Text, i == 0);
                rt.SetFromFont(FontName, runs[i].Size, runs[i].Bold);
                rt.Color = RunColors[i % RunColors.Length];
                rt.HorizontalAlignment = alignment;
            }

            return shape;
        }

        /// <summary>
        /// Returns one line per tspan in document order: x, y, the text's own width shaped in the
        /// tspan's font (px), where that text ends (x + width), and the text. When the next tspan
        /// on the same line starts before "end", the runs overlap.
        /// </summary>
        private static IEnumerable<string> GetTspans(OpenTypeFontEngine engine, string svg)
        {
            foreach (Match m in Regex.Matches(svg, "<tspan([^>]*)>([^<]*)</tspan>"))
            {
                var attributes = m.Groups[1].Value;
                var text = WebUtility.HtmlDecode(m.Groups[2].Value);
                var x = ParseNumber(GetAttribute(attributes, "x"));
                var y = ParseNumber(GetAttribute(attributes, "y"));
                var family = (GetAttribute(attributes, "font-family") ?? FontName).Split(',')[0].Trim().Trim('\'', '"');
                var sizePx = ParseNumber(GetAttribute(attributes, "font-size"));
                var bold = GetAttribute(attributes, "font-weight") == "bold";

                double widthPx = double.NaN;
                if (!double.IsNaN(sizePx))
                {
                    var shaper = engine.GetMeasurementShaper(family, bold ? FontSubFamily.Bold : FontSubFamily.Regular);
                    // Shaping at the pixel size gives the width in pixels.
                    widthPx = shaper.ShapeLight(text).GetWidthInPoints((float)sizePx);
                }

                yield return string.Format(CultureInfo.InvariantCulture,
                    "x={0,8:F2}  y={1,7:F2}  width={2,7:F2}  end={3,8:F2}  {4}{5} '{6}'",
                    x, y, widthPx, x + widthPx, family, bold ? " Bold" : "", text);
            }
        }

        private static string? GetAttribute(string attributes, string name)
        {
            // EPPlus writes some attributes with a space before '=' (x ="0").
            var m = Regex.Match(attributes, "(?:^|\\s)" + Regex.Escape(name) + "\\s*=\\s*\"([^\"]*)\"");
            return m.Success ? m.Groups[1].Value : null;
        }

        private static double ParseNumber(string? value)
        {
            if (string.IsNullOrEmpty(value))
                return double.NaN;
            double result;
            return double.TryParse(value.Replace("px", ""), NumberStyles.Float, CultureInfo.InvariantCulture, out result)
                ? result
                : double.NaN;
        }
    }
}
