using EPPlus.Fonts.OpenType;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style;
using System;
using System.Drawing;
using System.Text.RegularExpressions;

namespace EPPlus.DrawingRenderer.Tests.Fonts
{
    /// <summary>
    /// Base class for tests that export shapes and charts to svg: creates packages and drawings, and
    /// reads fonts, stacks and text runs from the svg.
    /// </summary>
    public abstract class SvgExportTestBase : FontTestBase
    {
        protected const string JapaneseText = "Sales あいう";

        protected OpenTypeFontEngine? _engine;

        protected ExcelPackage CreatePackage(Action<IEpplusFontConfiguration>? configure = null)
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

        protected static ExcelShape AddShape(ExcelPackage p, string text, string fontName, bool bold = false)
        {
            var ws = p.Workbook.Worksheets.Add("Sheet" + (p.Workbook.Worksheets.Count + 1));
            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetSize(300, 200);
            shape.TextBody.Paragraphs.Add(text).TextRuns[0].SetFromFont(fontName, 11, bold);
            return shape;
        }

        /// <summary>Hiragana is drawn with BIZ UDGothic, which an OnFontEmbedding callback may forbid.</summary>
        protected static void JapaneseFallback(IEpplusFontConfiguration cfg, bool embedFallback)
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

        protected static List<string> FontFaceFamilies(string svg)
        {
            return Regex.Matches(svg, "@font-face\\{font-family:'([^']*)'")
                .Cast<Match>().Select(m => m.Groups[1].Value).ToList();
        }

        protected static string[] Families(string fontFamilyAttribute)
        {
            return System.Net.WebUtility.HtmlDecode(fontFamilyAttribute)
                .Split(',').Select(f => f.Trim().Trim('\'')).ToArray();
        }

        /// <summary>The innermost tspans, the ones that hold text.</summary>
        protected static MatchCollection TextTspans(string svg)
        {
            return Regex.Matches(svg, "<tspan([^>]*)>([^<]+)</tspan>");
        }

        protected static string? Attribute(string attributes, string name)
        {
            var m = Regex.Match(attributes, "\\s" + Regex.Escape(name) + "=\"([^\"]*)\"");
            return m.Success ? m.Groups[1].Value : null;
        }

        /// <summary>The width in pixels, measured with the engine of the package the test created.</summary>
        protected double MeasurePixels(string text, string family, FontSubFamily subFamily, float size)
        {
            var points = _engine!.GetMeasurementShaper(family, subFamily).ShapeLight(text).GetWidthInPoints(size);
            return points / 72d * 96d;
        }

        // -----------------------------------------------------------------------------------------
        // Drawings with formatted text
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Adds a white rectangle with a gray border to a new worksheet, lets the caller add the text,
        /// and returns the svg.
        /// </summary>
        protected static string RenderFormattedShape(ExcelPackage p, int widthPixels, int heightPixels, Action<ExcelShape> addText)
        {
            var ws = p.Workbook.Worksheets.Add("FormattedShape");
            var shape = ws.Drawings.AddShape("FormattedShape", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 1, 0);
            shape.SetSize(widthPixels, heightPixels);
            shape.Fill.Style = eFillStyle.SolidFill;
            shape.Fill.Color = Color.White;
            shape.Border.Fill.Color = Color.Gray;
            addText(shape);
            return shape.ToSvg();
        }

        /// <summary>
        /// Adds a black 11 pt run in the given font to the shape.
        /// </summary>
        protected static ExcelParagraph AddRun(ExcelShape shape, string text, string font)
        {
            var run = shape.RichText.Add(text);
            run.LatinFont = font;
            run.Size = 11;
            run.Color = Color.Black;
            return run;
        }

        /// <summary>
        /// Renders a clustered column chart with a title and four categories.
        /// </summary>
        protected static string RenderColumnChart(ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets.Add("ColumnChart");
            ws.Cells["A1"].Value = "Quarter";
            ws.Cells["B1"].Value = "Sales";
            var quarters = new[] { "Q1", "Q2", "Q3", "Q4" };
            var values = new[] { 120d, 180d, 150d, 210d };
            for (int i = 0; i < quarters.Length; i++)
            {
                ws.Cells[i + 2, 1].Value = quarters[i];
                ws.Cells[i + 2, 2].Value = values[i];
            }

            var chart = ws.Drawings.AddChart("ColumnChart", eChartType.ColumnClustered);
            var serie = chart.Series.Add(ws.Cells["B2:B5"], ws.Cells["A2:A5"]);
            serie.Header = "Sales";
            chart.Title.Text = "Sales per quarter";
            chart.SetPosition(1, 0, 4, 0);
            chart.SetSize(480, 300);
            return chart.ToSvg();
        }
    }
}