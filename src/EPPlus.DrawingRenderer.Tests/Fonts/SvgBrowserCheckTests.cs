using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net;
using System.Text;

namespace EPPlus.DrawingRenderer.Tests.Fonts
{
    /// <summary>
    /// Step 7: generates a page for checking the svg export in browsers. Not a regression test.
    /// </summary>
    [TestClass]
    public class SvgBrowserCheckTests : SvgExportTestBase
    {
        // The folder under the test output folder (TestBase), written the way TestBase builds paths.
        private const string OutputFolder = "svg\\SvgBrowserCheck\\";

        /// <summary>
        /// Writes svg files and a page for checking the svg export in browsers to svg\SvgBrowserCheck
        /// under the test output folder (TestBase). Open index.html in Chrome, Firefox, Edge and Safari
        /// and go through the checks on the page. A screenshot from Excel named &lt;case&gt;.excel.png
        /// placed in the folder is shown next to its case the next time the page is written.
        /// Where files cannot be written, as in some build environments, the test is inconclusive.
        /// </summary>
        [TestMethod]
        [TestCategory("Manual")]
        public void WriteBrowserCheckPage()
        {
            RequireFont(TestFolderEngine, "Roboto");

            var html = new StringBuilder();
            AppendPageStart(html);

            foreach (var c in GetCases())
            {
                var svg = c.Render();
                // TestBase writes to the configured output folder and turns a failed write into Inconclusive.
                SaveTextFileToWorkbook(OutputFolder + c.File + ".svg", svg);
                AppendCaseRow(html, c, svg);
            }

            html.AppendLine("</table></body></html>");
            SaveTextFileToWorkbook(OutputFolder + "index.html", html.ToString());
            Console.WriteLine("Browser check page: " + GetOutputFile(OutputFolder, "index.html").FullName);
        }

        private List<BrowserCase> GetCases()
        {
            var robotoBold = TestFolderEngine.LoadFont("Roboto", FontSubFamily.Bold);
            var boldNote = robotoBold != null && robotoBold.NameTable.GetSubfamilyEnum() == FontSubFamily.Bold 
                ? 
                "Roboto has a bold file in the test folder, so this case checks that the bold face is used rather than synthesis."
                : "Roboto has no bold file in the test folder, so the browser synthesizes bold and italic.";

            const string longText = "The quick brown fox jumps over the lazy dog. Sphinx of black quartz, judge my vow. Office, affluent, waffle.";

            return new List<BrowserCase>
            {
                new BrowserCase("roboto-embedded", "Roboto, embedded",
                    "Drawn with the subset ABCDEF+Roboto. Ligatures in 'Office, affluent, waffle' if the font has them.",
                    () => { using (var p = CreatePackage()) { return AddShape(p, "Hello world 1 - Office, affluent, waffle", "Roboto").ToSvg(); } }),

                new BrowserCase("calibri-by-name", "Calibri, by name with safety net",
                    "Calibri may not be embedded. With Calibri installed (Windows) it is drawn with Calibri; without it (Mac, Linux) with the embedded safety net, normally Carlito or Archivo Narrow. The width is the same either way.",
                    () => { using (var p = CreatePackage()) { return AddShape(p, "Hello world 1 - Calibri", "Calibri").ToSvg(); } }),

                new BrowserCase("aptos-narrow-embedded", "Aptos Narrow, EmbedFonts = true",
                    "Measured as Aptos Narrow (from .fmtr metrics if it is not installed on the server). Written by name, followed by an embedded safety net. Browsers without Aptos Narrow use the safety net.",
                    () => { using (var p = CreatePackage()) { return AddShape(p, "Hello world 1 - Aptos Narrow", "Aptos Narrow").ToSvg(); } }),

                new BrowserCase("aptos-narrow-no-embed", "Aptos Narrow, EmbedFonts = false",
                    "Measured as Calibri (web substitution). No @font-face; the stack starts with 'Calibri'.",
                    () => { using (var p = CreatePackage()) { return AddShape(p, "Hello world 1 - Aptos Narrow", "Aptos Narrow").ToSvg(o => o.EmbedFonts = false); } }),

                new BrowserCase("roboto-no-embed", "Roboto, EmbedFonts = false",
                    "No @font-face. Drawn with Roboto only where it is installed, otherwise with a font from the stack; the width stays the same.",
                    () => { using (var p = CreatePackage()) { return AddShape(p, "Hello world 1", "Roboto").ToSvg(o => o.EmbedFonts = false); } }),

                new BrowserCase("japanese-fallback-by-name", "Japanese, glyph fallback by name",
                    "The run is split: 'Sales ' with the Roboto subset, the kana with BIZ UDGothic or the next font in the Japanese chain (Yu Gothic, Meiryo), never a Chinese font. Check the positions of both parts in Firefox and Safari.",
                    () => { using (var p = CreatePackage(cfg => JapaneseFallback(cfg, embedFallback: false))) { return AddShape(p, JapaneseText, "Roboto").ToSvg(); } }),

                new BrowserCase("japanese-fallback-embedded", "Japanese, glyph fallback embedded",
                    "One run, two subsets in the stack. The kana are drawn with the BIZ UDGothic subset.",
                    () => { using (var p = CreatePackage(cfg => JapaneseFallback(cfg, embedFallback: true))) { return AddShape(p, JapaneseText, "Roboto").ToSvg(); } }),

                new BrowserCase("bold-italic", "Bold and italic",
                    boldNote + " In the hostile column the page sets font-synthesis:none on every element; note whether synthesized bold or italic disappears there.",
                    () => { using (var p = CreatePackage()) { return RenderFormattedShape(p, 360, 80, s =>
                    {
                        AddRun(s, "Regular ", "Roboto");
                        AddRun(s, "Bold ", "Roboto").Bold = true;
                        AddRun(s, "Italic ", "Roboto").Italic = true;
                        var both = AddRun(s, "Bold italic", "Roboto");
                        both.Bold = true;
                        both.Italic = true;
                    }); } }),

                new BrowserCase("decoration", "Underline and strike-through",
                    "'Both' has both an underline and a strike-through. 'Red line' has a red underline under black text. Both must remain in the hostile column, which sets text-decoration:none.",
                    () => { using (var p = CreatePackage()) { return RenderFormattedShape(p, 360, 80, s =>
                    {
                        var both = AddRun(s, "Both", "Roboto");
                        both.UnderLine = eUnderLineType.Single;
                        both.Strike = eStrikeType.Single;
                        AddRun(s, " and ", "Roboto");
                        var red = AddRun(s, "Red line", "Roboto");
                        red.UnderLine = eUnderLineType.Single;
                        red.UnderLineColor = Color.Red;
                    }); } }),

                new BrowserCase("superscript", "Superscript and subscript",
                    "The 2 is raised and the 2 in H2O is lowered. Known: both move about 4/3 too far (dy converted twice; on the residue list).",
                    () => { using (var p = CreatePackage()) { return RenderFormattedShape(p, 360, 80, s =>
                    {
                        AddRun(s, "E = mc", "Roboto");
                        AddRun(s, "2", "Roboto").Baseline = 30;
                        AddRun(s, "   H", "Roboto");
                        AddRun(s, "2", "Roboto").Baseline = -25;
                        AddRun(s, "O", "Roboto");
                    }); } }),

                new BrowserCase("emoji", "Emoji and symbols",
                    "The emoji and the check mark are drawn with the embedded Noto Emoji subset (monochrome), not with a color emoji font of the system.",
                    () => { using (var p = CreatePackage()) { return AddShape(p, "Smile \U0001F642 done \u2713", "Roboto").ToSvg(); } }),

                new BrowserCase("wrapping", "Line breaks",
                    "The lines break after the same words as in Excel, and no line is wider than the shape.",
                    () => { using (var p = CreatePackage()) { return RenderFormattedShape(p, 180, 140, s => AddRun(s, longText, "Roboto")); } }),

                new BrowserCase("chart", "Chart",
                    "Title, axis labels and values are placed and sized as in Excel. With several svgs on the page, gradients or patterns of one svg must not show up in another (ids are not unique across svgs).",
                    () => { using (var p = CreatePackage()) { return RenderColumnChart(p); } }),
            };
        }

        private static void AppendPageStart(StringBuilder html)
        {
            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html lang=\"en\"><head><meta charset=\"utf-8\"><title>EPPlus SVG browser check</title>");
            html.AppendLine("<style>");
            html.AppendLine("body{font-family:system-ui,sans-serif;margin:16px;color:#222;background:#fff}");
            html.AppendLine("table{border-collapse:collapse}td,th{border:1px solid #ccc;padding:8px;vertical-align:top}th{background:#f4f4f4;text-align:left}");
            html.AppendLine(".check{font-size:13px;max-width:280px}.missing{color:#999;font-size:12px}");
            // A hostile host page: rules that reach text and tspan directly, not only through inheritance.
            html.AppendLine(".hostile svg,.hostile svg *{font-family:'Comic Sans MS',cursive,serif;font-size:20px;font-weight:normal;font-style:normal;" +
                "letter-spacing:2px;word-spacing:6px;font-variant-ligatures:none;font-kerning:none;text-rendering:optimizeSpeed;" +
                "font-synthesis:none;text-decoration:none}");
            html.AppendLine("</style></head><body>");
            html.AppendLine("<h1>EPPlus SVG browser check</h1>");
            html.Append("<p>Generated ").Append(WebUtility.HtmlEncode(DateTime.Now.ToString("yyyy-MM-dd HH:mm")))
                .AppendLine(". Open in Chrome, Firefox, Edge and Safari.</p>");
            html.AppendLine("<h2>General checks</h2><ul>");
            html.AppendLine("<li>The img, inline and hostile columns show the same text in the same places, and the text fits the shape as in Excel.</li>");
            html.AppendLine("<li>Hostile column: font, size, weight and style must not change (inline style). Letter and word spacing must not change the width of a run (textLength).</li>");
            html.AppendLine("<li>Which font a run is drawn with: Chrome/Edge DevTools, Elements, Computed, Rendered Fonts. Firefox: Inspector, Fonts.</li>");
            html.AppendLine("<li>Zoom to 90, 110 and 125 percent: text scales smoothly and keeps its width relative to the shape (text-rendering:geometricPrecision, Chrome and Firefox).</li>");
            html.AppendLine("<li>The open link shows the standalone svg; it looks the same as the img column.</li>");
            html.AppendLine("</ul>");
            html.AppendLine("<table><tr><th>Case</th><th>img</th><th>inline</th><th>inline, hostile css</th><th>Excel</th><th>File</th></tr>");
        }

        private static void AppendCaseRow(StringBuilder html, BrowserCase c, string svg)
        {
            var excel = c.File + ".excel.png";
            html.AppendLine("<tr>");
            html.Append("<td class=\"check\"><b>").Append(WebUtility.HtmlEncode(c.Title)).Append("</b><br>")
                .Append(WebUtility.HtmlEncode(c.Check)).AppendLine("</td>");
            html.Append("<td><img src=\"").Append(c.File).AppendLine(".svg\" alt=\"\"></td>");
            html.Append("<td>").Append(svg).AppendLine("</td>");
            html.Append("<td class=\"hostile\">").Append(svg).AppendLine("</td>");
            html.Append("<td>");
            if (GetOutputFile(OutputFolder, excel).Exists)
                html.Append("<img src=\"").Append(excel).Append("\" alt=\"\">");
            else
                html.Append("<span class=\"missing\">").Append(excel).Append("</span>");
            html.AppendLine("</td>");
            html.Append("<td><a href=\"").Append(c.File).AppendLine(".svg\">open</a></td>");
            html.AppendLine("</tr>");
        }

        private sealed class BrowserCase
        {
            internal BrowserCase(string file, string title, string check, Func<string> render)
            {
                File = file;
                Title = title;
                Check = check;
                Render = render;
            }

            /// <summary>The file name of the svg, without extension.</summary>
            internal string File { get; }
            internal string Title { get; }

            /// <summary>What to look for in the browser.</summary>
            internal string Check { get; }
            internal Func<string> Render { get; }
        }
    }
}