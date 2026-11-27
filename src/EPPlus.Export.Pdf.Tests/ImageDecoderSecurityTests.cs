using EPPlus.Export.Pdf.Helpers;          // internal PngDecoder (InternalsVisibleTo is already set)
using EPPlus.Export.Pdf.Tests;            // PdfTestBase -> TestBase (license + OpenTemplatePackage)
using OfficeOpenXml;
using OfficeOpenXml.Export.PdfExport;     // SaveAsPdf extension methods
using System;
using System.IO;
using System.Text;

namespace EPPlusTest.PDF
{
    /// <summary>
    /// Security regression tests for the embedded-image decoders. The attack: a tiny PNG whose
    /// IHDR *declares* enormous dimensions (30000x30000) while the IDAT is a real 1x1 pixel.
    /// The decoder sizes its output buffer from the header (width*height*3 in Int32), which
    /// overflows to a negative length and throws, aborting the entire export. Reachable by anyone
    /// who can get an image into a workbook that gets converted to PDF.
    /// </summary>
    [TestClass]
    public class ImageDecoderSecurityTests : PdfTestBase
    {
        // IHDR says 30000x30000, 8-bit, colour type 6 (RGBA). IDAT is a valid zlib stream for one
        // 1x1 pixel. (Exact bytes of evil_oversized_30000x30000.png.)
        private static readonly byte[] OversizedRgbaPng = new byte[]
        {
            137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13,
            73, 72, 68, 82, 0, 0, 117, 48, 0, 0, 117, 48,
            8, 6, 0, 0, 0, 102, 39, 248, 186, 0, 0, 0,
            13, 73, 68, 65, 84, 120, 218, 99, 248, 207, 192, 240,
            31, 0, 5, 0, 1, 255, 86, 199, 47, 13, 0, 0,
            0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
        };

        private const int OversizedW = 30000;
        private const int OversizedH = 30000;
        private const int RgbaColorType = 6;

        /// <summary>The header parses as a PNG but the dimensions must be rejected by the guard.</summary>
        [TestMethod]
        public void PngDecoder_RejectsOversizedDimensions()
        {
            Assert.IsTrue(PngDecoder.IsPng(OversizedRgbaPng));

            bool ok = PngDecoder.ReadPngHeader(
                OversizedRgbaPng, out int w, out int h, out int bitDepth, out int colorType, out _);

            Assert.IsFalse(ok, "Oversized PNG header must be rejected by the ImageLimits guard.");
            // Dimensions are still parsed out (the guard runs after reading them) — confirm the
            // payload really is the oversized one, so this test can't pass for the wrong reason.
            Assert.AreEqual(OversizedW, w);
            Assert.AreEqual(OversizedH, h);
            Assert.AreEqual(RgbaColorType, colorType);
        }

        /// <summary>Calling the decode path directly must throw a clean, typed error — never overflow/OOM.</summary>
        [TestMethod]
        public void DecodePngWithAlpha_OnOversized_ThrowsInvalidData()
        {
            Assert.ThrowsExactly<InvalidDataException>(() =>
                PngDecoder.DecodePngWithAlpha(
                    OversizedRgbaPng, OversizedW, OversizedH,
                    blackAndWhite: false, colorType: RgbaColorType,
                    out _, out _));
        }

        /// <summary>End-to-end: a workbook containing the oversized image still exports cleanly.</summary>
        [TestMethod]
        public void Export_WithOversizedPng_DoesNotThrow_AndProducesValidPdf()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var ws = p.Workbook.Worksheets[0];

            using (var imgStream = new MemoryStream(OversizedRgbaPng))
                ws.Drawings.AddPicture("evil_oversized", imgStream);

            using var ms = new MemoryStream();

            ws.SaveAsPdf(ms);   // must NOT throw — the bad image is skipped/blanked

            byte[] bytes = ms.ToArray();
            Assert.IsTrue(bytes.Length > 8, "No PDF was produced.");
            Assert.AreEqual("%PDF-", Encoding.ASCII.GetString(bytes, 0, 5), "Missing PDF header.");

            string tail = Encoding.ASCII.GetString(bytes, Math.Max(0, bytes.Length - 8), Math.Min(8, bytes.Length));
            Assert.IsTrue(tail.Contains("%%EOF"), "Missing %%EOF trailer — document is truncated.");
        }
    }
}