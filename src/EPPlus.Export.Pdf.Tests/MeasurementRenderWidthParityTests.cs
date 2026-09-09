/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/07/2026         EPPlus Software AB           Measurement/render width parity (WP1)
 *************************************************************************************************/
using EPPlus.Export.Pdf.Settings;
using EPPlus.Export.Pdf.Tests;
using EPPlus.Fonts.OpenType;
using OfficeOpenXml;
using OfficeOpenXml.Export.PdfExport;
using OfficeOpenXml.Packaging.Ionic.Zlib;
using OfficeOpenXml.Style;
using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace EPPlusTest.PDF
{
    /// <summary>
    /// WP1's last, previously unverified acceptance criterion: the measurement path
    /// (PdfCatalog.GetCellCollectionFromRange, which shapes against the FULL font and skips
    /// BuildSubsets) and the render path (PdfCatalog.Save, which shapes against the SUBSETTED
    /// font) must produce the same width for the same string and font.
    ///
    /// This matters because autofit and row-height calculations are driven by the measurement
    /// path's PdfCell.TotalTextLength, while what actually appears on the page is driven by the
    /// render path. If the two ever diverged, a cell could measure as fitting a string, but the
    /// rendered glyphs would run wider (or narrower) than that measurement said - text clipped or
    /// falsely wrapped, or autofit-shrunk column widths leaving unnecessary blank margin.
    ///
    /// The test:
    ///   1. Gets the MEASURED width via GetCellCollectionFromRange -> PdfCell.TotalTextLength.
    ///   2. Renders the same sheet to actual PDF bytes.
    ///   3. Independently recomputes the RENDERED width by parsing the content stream's TJ array
    ///      and looking up each glyph's advance in the embedded (subsetted) font's own hmtx
    ///      table, loaded via this same library's OpenTypeFontFactory - not via any external
    ///      tool, and not by trusting the PDF's own /W CID-widths array, since that would only be
    ///      checking the writer's arithmetic against itself.
    ///   4. Asserts the two match within a small tolerance.
    ///
    /// The tolerance is not slack for a real discrepancy: PdfContentStream writes kerning
    /// adjustments through ToPdfStringF0 (rounded to the nearest integer TJ unit), while
    /// TotalTextLength is computed from the exact, unrounded kerning value. That integer rounding
    /// is the only expected source of difference, and it is at most a few hundredths of a point
    /// per kerned pair.
    ///
    /// Uses "AVATAR Wa To" in Roboto - a repo test font with real kerning and no accents - so the
    /// content stream only contains glyph tokens and kerning numbers, with no Ts (mark offset) or
    /// mid-string font switches to account for.
    /// </summary>
    [TestClass]
    public class MeasurementRenderWidthParityTests : PdfTestBase
    {
        // Encoding.Latin1 (a 1:1 byte<->char mapping, needed so string indices from Regex
        // line up with byte offsets in the original array) is only available from .NET 5+.
        // GetEncoding("ISO-8859-1") gives the same mapping on both net481 and net8.0.
        private static readonly Encoding Latin1 = Encoding.GetEncoding("ISO-8859-1");

        private const string ReproFontFamily = "Roboto";
        private const float ReproFontSize = 48f;
        private const string ReproText = "AVATAR Wa To";

        // PdfContentStream.ToPdfStringF0 rounds each kerning TJ number to the nearest integer
        // (1/1000 em unit). At 48pt that is at most 0.048pt of rounding error per kerned pair;
        // "AVATAR Wa To" has six kerned pairs (A+V, V+A, A+T, T+A, W+a, T+o), so 0.5pt covers the
        // worst case with headroom.
        private const double ToleranceInPoints = 0.5;

        [TestMethod]
        public void MeasuredWidth_MatchesRenderedWidth_ForKernedText()
        {
            using (var package = OpenPackage("WidthParity.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("Parity");
                sheet.Cells["A1"].Value = ReproText;
                sheet.Cells["A1"].Style.Font.Name = ReproFontFamily;
                sheet.Cells["A1"].Style.Font.Size = ReproFontSize;
                sheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                sheet.Cells["A1"].Style.Indent = 0;
                sheet.Column(1).Width = 60;
                sheet.Row(1).CustomHeight = true;
                sheet.Row(1).Height = 70;

                var engine = package.Workbook.RenderContext.FontEngine;
                var settings = new PdfPageSettings(engine);
                var range = sheet.Cells["A1"];

                double measuredWidth = GetMeasuredWidth(settings, range);

                var pdfBytes = RenderToBytes(settings, sheet);

                // Written out unconditionally (not just on failure) so a parsing mismatch can be
                // diagnosed from the actual bytes rather than guessed at from source alone.
                File.WriteAllBytes(Path.Combine(_pdfPath, "WidthParity.pdf"), pdfBytes);

                double renderedWidth = GetRenderedWidth(pdfBytes, ReproFontSize);

                Assert.AreNotEqual(0, measuredWidth, "measured width must not be zero - shaping did not run");
                Assert.AreNotEqual(0, renderedWidth, "rendered width must not be zero - TJ parsing found nothing");

                Assert.AreEqual(
                    measuredWidth,
                    renderedWidth,
                    ToleranceInPoints,
                    $"measurement path width ({measuredWidth:F3}pt) and render path width "
                    + $"({renderedWidth:F3}pt) must match within rounding. A real divergence here "
                    + "means the measurement path (full font) and render path (subsetted font) "
                    + "are shaping this string differently.");
            }
        }

        private static double GetMeasuredWidth(PdfPageSettings settings, ExcelRangeBase range)
        {
            var catalog = new PdfCatalog();
            var cells = catalog.GetCellCollectionFromRange(settings, range);
            return cells[range.Start.Row, range.Start.Column].TotalTextLength;
        }

        private static byte[] RenderToBytes(PdfPageSettings settings, ExcelWorksheet sheet)
        {
            using (var stream = new MemoryStream())
            {
                new PdfCatalog(settings, sheet).Save(stream);
                return stream.ToArray();
            }
        }

        private static double GetRenderedWidth(byte[] pdfBytes, float fontSize)
        {
            string pdfText = Latin1.GetString(pdfBytes);

            byte[] fontBytes = ExtractFirstFontFile2(pdfBytes, pdfText);
            Assert.IsNotNull(fontBytes, "no /FontFile2 found in the rendered PDF");

            var font = OpenTypeFontFactory.CreateFromBytes(fontBytes);
            ushort unitsPerEm = font.HeadTable.UnitsPerEm;

            string contentStream = ExtractContentStreamWithTJ(pdfBytes, pdfText);
            Assert.IsNotNull(contentStream, "no content stream containing a TJ array was found");

            var tjMatch = Regex.Match(contentStream, @"\[(?<body>.*?)\]\s*TJ", RegexOptions.Singleline);
            Assert.IsTrue(tjMatch.Success, "no [...] TJ array found in the content stream");

            double widthInFontUnits = 0;

            foreach (Match token in Regex.Matches(
                tjMatch.Groups["body"].Value,
                @"<(?<glyph>[0-9A-Fa-f]{4})>|(?<num>-?\d+(\.\d+)?)"))
            {
                if (token.Groups["glyph"].Success)
                {
                    ushort glyphId = ushort.Parse(token.Groups["glyph"].Value, NumberStyles.HexNumber);
                    widthInFontUnits += font.HmtxTable.GetAdvanceWidth(glyphId);
                }
                else if (token.Groups["num"].Success)
                {
                    // TJ numbers are in 1/1000 text space units already, independent of the
                    // font's own unitsPerEm - convert to the same font-unit space as the glyph
                    // advances above so both can be summed together before the final scale to
                    // points.
                    double tjNumber = double.Parse(token.Groups["num"].Value, CultureInfo.InvariantCulture);
                    widthInFontUnits -= tjNumber / 1000.0 * unitsPerEm;
                }
            }

            return widthInFontUnits / unitsPerEm * fontSize;
        }

        /// <summary>
        /// Finds the first "N 0 obj ... stream ... endstream" block that contains "TJ" once
        /// inflated, skipping font-file streams. EPPlus's PDF writer never emits object or
        /// cross-reference streams, so every stream is either an uncompressed page content
        /// stream candidate or a FlateDecode-compressed one - this handles both.
        /// </summary>
        private static string ExtractContentStreamWithTJ(byte[] pdfBytes, string pdfText)
        {
            foreach (Match m in Regex.Matches(pdfText, @"<<(?<dict>.*?)>>\s*stream\r?\n", RegexOptions.Singleline))
            {
                if (m.Groups["dict"].Value.Contains("FontFile"))
                    continue;

                int start = m.Index + m.Length;
                int end = pdfText.IndexOf("endstream", start, StringComparison.Ordinal);
                if (end < 0) continue;

                byte[] raw = Latin1Slice(pdfBytes, start, end);
                byte[] data = m.Groups["dict"].Value.Contains("FlateDecode") ? Inflate(raw) : raw;

                string text = Latin1.GetString(data);
                if (text.Contains("TJ"))
                    return text;
            }

            return null;
        }

        /// <summary>
        /// Finds the object referenced by the first /FontFile2 entry and returns its
        /// (decompressed) stream bytes - the embedded, SUBSETTED font, which is exactly what the
        /// render path actually shaped against and what a PDF viewer would use.
        /// </summary>
        private static byte[] ExtractFirstFontFile2(byte[] pdfBytes, string pdfText)
        {
            var refMatch = Regex.Match(pdfText, @"/FontFile2\s+(?<num>\d+)\s+0\s+R");
            if (!refMatch.Success) return null;

            string objNum = refMatch.Groups["num"].Value;

            var objMatch = Regex.Match(pdfText, $@"(?<!\d){objNum}\s+0\s+obj\s*(?<dict><<.*?>>)\s*stream\r?\n", RegexOptions.Singleline);
            if (!objMatch.Success) return null;

            int start = objMatch.Index + objMatch.Length;
            int end = pdfText.IndexOf("endstream", start, StringComparison.Ordinal);
            if (end < 0) return null;

            byte[] raw = Latin1Slice(pdfBytes, start, end);
            return objMatch.Groups["dict"].Value.Contains("FlateDecode") ? Inflate(raw) : raw;
        }

        /// <summary>
        /// Slices the ORIGINAL byte array using offsets found via a Latin-1 string view of it.
        /// Latin-1 is a 1:1 byte<->char mapping, so string indices returned by Regex against the
        /// Latin-1-decoded text correspond exactly to byte offsets in the original array - unlike
        /// UTF-8 or any other multi-byte encoding, which would desynchronize the two.
        /// </summary>
        private static byte[] Latin1Slice(byte[] bytes, int start, int end)
        {
            var result = new byte[end - start];
            Array.Copy(bytes, start, result, 0, result.Length);
            return result;
        }

        private static byte[] Inflate(byte[] flateCompressed)
        {
            using (var input = new MemoryStream(flateCompressed))
            using (var zlib = new ZlibStream(input, CompressionMode.Decompress))
            using (var output = new MemoryStream())
            {
                zlib.CopyTo(output);
                return output.ToArray();
            }
        }
    }
}