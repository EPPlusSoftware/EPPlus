/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using EPPlus.Export.Pdf.Settings;
using EPPlus.Export.Pdf.Settings.PdfPageSizes;
using EPPlus.Export.Pdf.Tests;
using OfficeOpenXml;
using OfficeOpenXml.Export.PdfExport;
using OfficeOpenXml.Export.PdfExport.Data;
using OfficeOpenXml.Export.PdfExport.Layout;
using OfficeOpenXml.Export.PdfExport.Settings;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Reflection.Metadata;
using System.Text;
using System.Text.RegularExpressions;

namespace EPPlusTest.PDF
{
    [TestClass]
    public class PdfTests : PdfTestBase
    {
        private static void AssertLooksLikePdf(byte[] bytes)
        {
            Assert.IsTrue(bytes.Length > 0, "PDF output is empty.");
            string head = Encoding.ASCII.GetString(bytes, 0, Math.Min(8, bytes.Length));
            Assert.IsTrue(head.StartsWith("%PDF-"), $"Missing PDF header. Got: '{head}'");
            int tailLen = Math.Min(8, bytes.Length);
            string tail = Encoding.ASCII.GetString(bytes, bytes.Length - tailLen, tailLen);
            Assert.IsTrue(tail.Contains("%%EOF"), "Missing %%EOF trailer marker.");
        }

        private static long ParseStartXref(byte[] bytes, int pdfStart)
        {
            string text = Encoding.ASCII.GetString(bytes, pdfStart, bytes.Length - pdfStart);
            int idx = text.LastIndexOf("startxref", StringComparison.Ordinal);
            Assert.IsTrue(idx >= 0, "startxref keyword not found.");
            int i = idx + "startxref".Length;
            while (i < text.Length && (text[i] == '\n' || text[i] == '\r' || text[i] == ' ')) i++;
            int start = i;
            while (i < text.Length && char.IsDigit(text[i])) i++;
            return long.Parse(text.Substring(start, i - start));
        }

        [TestMethod]
        public void SaveWorksheetAsPdfTest1()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var ws = p.Workbook.Worksheets[0];
            string path = _pdfPath + "WorksheetTest1.pdf";
            ws.SaveAsPdf(path);
            Assert.IsTrue(File.Exists(path), "PDF file was not created.");
            AssertLooksLikePdf(File.ReadAllBytes(path));
        }

        [TestMethod]
        public void SaveWorksheetAsPdfTest2()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var ws = p.Workbook.Worksheets[0];
            ws.PrinterSettings.Orientation = eOrientation.Landscape;
            ws.PrinterSettings.ShowGridLines = false;
            ws.PrinterSettings.ShowHeaders = false;
            ws.PrinterSettings.PaperSize = ePaperSize.A3;
            string path = _pdfPath + "WorksheetTest2.pdf";
            ws.SaveAsPdf(path);
            Assert.IsTrue(File.Exists(path), "PDF file was not created.");
            AssertLooksLikePdf(File.ReadAllBytes(path));
        }

        [TestMethod]
        public void SaveRangeAsPdfTest1()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var range = p.Workbook.Worksheets[0].Cells["D3:F6"];
            string path = _pdfPath + "RangeTest1.pdf";
            range.SaveAsPdf(path);
            Assert.IsTrue(File.Exists(path), "PDF file was not created.");
            AssertLooksLikePdf(File.ReadAllBytes(path));
        }

        [TestMethod]
        public void SaveWorkbookAsPdfTest1()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            string path = _pdfPath + "WorkbookTest1.pdf";
            wb.SaveAsPdf(path);
            Assert.IsTrue(File.Exists(path), "PDF file was not created.");
            AssertLooksLikePdf(File.ReadAllBytes(path));
        }

        [TestMethod]
        public void SaveWorksheetsAsPdfTest2()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            var ws0 = wb.Worksheets[0];
            var ws1 = wb.Worksheets[1];
            var ws2 = wb.Worksheets[2];
            string path = _pdfPath + "WorksheetsTest2.pdf";
            wb.SaveAsPdf(path, ws0, ws1, ws2);
            Assert.IsTrue(File.Exists(path), "PDF file was not created.");
            AssertLooksLikePdf(File.ReadAllBytes(path));
        }

        [TestMethod]
        public void SaveWorksheetsAsPdfTest1()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            var ws0 = wb.Worksheets[0];
            var ws2 = wb.Worksheets[2];
            string path = _pdfPath + "WorksheetsTest1.pdf";
            wb.SaveAsPdf(path, ws0, ws2);
            Assert.IsTrue(File.Exists(path), "PDF file was not created.");
            AssertLooksLikePdf(File.ReadAllBytes(path));
        }

        [TestMethod]
        public void SaveRangesAsPdfTest1()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            var ws = wb.Worksheets[0];
            var r1 = ws.Cells["D3:F6"];
            var r2 = ws.Cells["B36:F39"];
            var r3 = ws.Cells["K49:Q58"];
            var r4 = ws.Cells["L142:Q147"];
            string path = _pdfPath + "RangesTest1.pdf";
            wb.SaveAsPdf(path, r1, r2, r3, r4);
            Assert.IsTrue(File.Exists(path), "PDF file was not created.");
            AssertLooksLikePdf(File.ReadAllBytes(path));
        }

        [TestMethod]
        public async Task SaveWorkbookAsPdfAsyncTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            string tempFile = Path.GetTempFileName();
            try
            {
                await wb.SaveAsPdfAsync(tempFile);
                AssertLooksLikePdf(File.ReadAllBytes(tempFile));
            }
            finally
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }
        }

        [TestMethod]
        public async Task SaveWorksheetsAsPdfAsyncTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            var ws0 = wb.Worksheets[0];
            var ws2 = wb.Worksheets[2];
            string tempFile = Path.GetTempFileName();
            try
            {
                await wb.SaveAsPdfAsync(tempFile, ws0, ws2);
                AssertLooksLikePdf(File.ReadAllBytes(tempFile));
            }
            finally
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }
        }

        [TestMethod]
        public async Task SaveWorksheetsAsPdfAsyncWithTokenTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            var ws0 = wb.Worksheets[0];
            var ws2 = wb.Worksheets[2];
            string tempFile = Path.GetTempFileName();
            try
            {
                await wb.SaveAsPdfAsync(tempFile, CancellationToken.None, ws0, ws2);
                AssertLooksLikePdf(File.ReadAllBytes(tempFile));
            }
            finally
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }
        }

        [TestMethod]
        public async Task SaveRangesAsPdfAsyncTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            var ws = wb.Worksheets[0];
            var r1 = ws.Cells["D3:F6"];
            var r2 = ws.Cells["B36:F39"];
            string tempFile = Path.GetTempFileName();
            try
            {
                await wb.SaveAsPdfAsync(tempFile, r1, r2);
                AssertLooksLikePdf(File.ReadAllBytes(tempFile));
            }
            finally
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }
        }

        [TestMethod]
        public async Task SaveRangesAsPdfAsyncWithTokenTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            var ws = wb.Worksheets[0];
            var r1 = ws.Cells["D3:F6"];
            var r2 = ws.Cells["B36:F39"];
            string tempFile = Path.GetTempFileName();
            try
            {
                await wb.SaveAsPdfAsync(tempFile, CancellationToken.None, r1, r2);
                AssertLooksLikePdf(File.ReadAllBytes(tempFile));
            }
            finally
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }
        }

        [TestMethod]
        public async Task SaveWorkbookAsPdfAsyncWithCanceledTokenThrowsTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            string tempFile = Path.GetTempFileName();
            try
            {
                using var cts = new CancellationTokenSource();
                cts.Cancel();

                await Assert.ThrowsExactlyAsync<TaskCanceledException>(
                    () => wb.SaveAsPdfAsync(tempFile, cts.Token));
            }
            finally
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }
        }

        [TestMethod]
        public void SaveWorkbookAsPdfToStreamTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            using var ms = new MemoryStream();
            wb.SaveAsPdf(ms);
            AssertLooksLikePdf(ms.ToArray());
        }

        [TestMethod]
        public void SaveWorksheetsAsPdfToStreamTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            var ws0 = wb.Worksheets[0];
            var ws2 = wb.Worksheets[2];
            using var ms = new MemoryStream();
            wb.SaveAsPdf(ms, ws0, ws2);
            AssertLooksLikePdf(ms.ToArray());
        }

        [TestMethod]
        public void SaveRangesAsPdfToStreamTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            var ws = wb.Worksheets[0];
            var r1 = ws.Cells["D3:F6"];
            var r2 = ws.Cells["B36:F39"];
            using var ms = new MemoryStream();
            wb.SaveAsPdf(ms, r1, r2);
            AssertLooksLikePdf(ms.ToArray());
        }

        [TestMethod]
        public async Task SaveWorkbookAsPdfToStreamAsyncTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            using var ms = new MemoryStream();
            await wb.SaveAsPdfAsync(ms);
            AssertLooksLikePdf(ms.ToArray());
        }

        [TestMethod]
        public async Task SaveAsPdfAsyncWithCanceledTokenThrowsTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            using var ms = new MemoryStream();
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            await Assert.ThrowsExactlyAsync<TaskCanceledException>(() => wb.SaveAsPdfAsync(ms, cts.Token));
        }

        [TestMethod]
        public void SaveAsPdfToStreamLeavesStreamOpenTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            using var ms = new MemoryStream();
            wb.SaveAsPdf(ms);
            Assert.IsTrue(ms.CanWrite, "Stream was closed by the export.");
            Assert.IsTrue(ms.CanRead, "Stream was closed by the export.");
            Assert.IsTrue(ms.Length > 0, "Nothing was written to the stream.");
        }

        [TestMethod]
        public void StreamOffsetsAreRelativeToPdfStartTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            using var ms = new MemoryStream();
            // Pre-fill the stream so the PDF does not start at offset 0.
            byte[] preamble = Encoding.ASCII.GetBytes("LEADING BYTES THAT ARE NOT PART OF THE PDF");
            ms.Write(preamble, 0, preamble.Length);
            int pdfStart = (int)ms.Position;
            wb.SaveAsPdf(ms);
            byte[] all = ms.ToArray();
            // The PDF itself still starts with the header at the captured position.
            string header = Encoding.ASCII.GetString(all, pdfStart, 5);
            Assert.AreEqual("%PDF-", header, "PDF was not written at the stream's current position.");
            // startxref must point at the xref table relative to the PDF start.
            long startXref = ParseStartXref(all, pdfStart);
            string atOffset = Encoding.ASCII.GetString(all, pdfStart + (int)startXref, 4);
            Assert.AreEqual("xref", atOffset, "startxref offset is not relative to the PDF start.");
        }

        [TestMethod]
        public void FileAndStreamProduceSamePdfLengthTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            string tempFile = Path.GetTempFileName();
            try
            {
                wb.SaveAsPdf(tempFile);
                long fileLength = new FileInfo(tempFile).Length;
                using var ms = new MemoryStream();
                wb.SaveAsPdf(ms);
                Assert.AreEqual(fileLength, ms.Length, "Stream output length differs from file output length.");
            }
            finally
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }
        }

        [TestMethod]
        public void SaveWorksheetToStreamViaCatalogTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var pageSettings = new PdfPageSettings(ws.Workbook.RenderContext.FontEngine);
            using var ms = new MemoryStream();
            var pdfCatalog = new PdfCatalog(pageSettings, ws);
            pdfCatalog.Save(ms);
            AssertLooksLikePdf(ms.ToArray());
        }

        [TestMethod]
        public void SaveAsPdfToNonWritableStreamThrowsTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            using var readOnly = new MemoryStream(new byte[16], writable: false);
            Assert.ThrowsExactly<ArgumentException>(() => wb.SaveAsPdf(readOnly));
        }

        [TestMethod]
        public void SaveWorksheetAsPdfToStreamTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var ws = p.Workbook.Worksheets[0];
            using var ms = new MemoryStream();
            ws.SaveAsPdf(ms);
            AssertLooksLikePdf(ms.ToArray());
        }

        [TestMethod]
        public async Task SaveWorksheetAsPdfToStreamAsyncTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var ws = p.Workbook.Worksheets[0];
            using var ms = new MemoryStream();
            await ws.SaveAsPdfAsync(ms);
            AssertLooksLikePdf(ms.ToArray());
        }

        [TestMethod]
        public void SaveWorksheetAsPdfToStreamLeavesStreamOpenTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var ws = p.Workbook.Worksheets[0];
            using var ms = new MemoryStream();
            ws.SaveAsPdf(ms);
            Assert.IsTrue(ms.CanWrite, "Stream was closed by the export.");
            Assert.IsTrue(ms.CanRead, "Stream was closed by the export.");
            Assert.IsTrue(ms.Length > 0, "Nothing was written to the stream.");
        }

        [TestMethod]
        public async Task SaveWorksheetAsPdfAsyncToStreamWithCanceledTokenThrowsTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var ws = p.Workbook.Worksheets[0];
            using var ms = new MemoryStream();
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            await Assert.ThrowsExactlyAsync<TaskCanceledException>(() => ws.SaveAsPdfAsync(ms, cts.Token));
        }

        [TestMethod]
        public void SaveWorksheetToStreamOffsetsAreRelativeToPdfStartTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var ws = p.Workbook.Worksheets[0];
            using var ms = new MemoryStream();
            byte[] preamble = Encoding.ASCII.GetBytes("LEADING BYTES THAT ARE NOT PART OF THE PDF");
            ms.Write(preamble, 0, preamble.Length);
            int pdfStart = (int)ms.Position;
            ws.SaveAsPdf(ms);
            byte[] all = ms.ToArray();
            string header = Encoding.ASCII.GetString(all, pdfStart, 5);
            Assert.AreEqual("%PDF-", header, "PDF was not written at the stream's current position.");
            long startXref = ParseStartXref(all, pdfStart);
            string atOffset = Encoding.ASCII.GetString(all, pdfStart + (int)startXref, 4);
            Assert.AreEqual("xref", atOffset, "startxref offset is not relative to the PDF start.");
        }

        [TestMethod]
        public void SaveWorksheetToNonWritableStreamThrowsTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var ws = p.Workbook.Worksheets[0];
            using var readOnly = new MemoryStream(new byte[16], writable: false);
            Assert.ThrowsExactly<ArgumentException>(() => ws.SaveAsPdf(readOnly));
        }

        [TestMethod]
        public void SaveRangeAsPdfToStreamTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var range = p.Workbook.Worksheets[0].Cells["D3:F6"];
            using var ms = new MemoryStream();
            range.SaveAsPdf(ms);
            AssertLooksLikePdf(ms.ToArray());
        }

        [TestMethod]
        public async Task SaveRangeAsPdfToStreamAsyncTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var range = p.Workbook.Worksheets[0].Cells["D3:F6"];
            using var ms = new MemoryStream();
            await range.SaveAsPdfAsync(ms);
            AssertLooksLikePdf(ms.ToArray());
        }

        [TestMethod]
        public void SaveRangeAsPdfToStreamLeavesStreamOpenTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var range = p.Workbook.Worksheets[0].Cells["D3:F6"];
            using var ms = new MemoryStream();
            range.SaveAsPdf(ms);
            Assert.IsTrue(ms.CanWrite, "Stream was closed by the export.");
            Assert.IsTrue(ms.CanRead, "Stream was closed by the export.");
            Assert.IsTrue(ms.Length > 0, "Nothing was written to the stream.");
        }

        [TestMethod]
        public async Task SaveRangeAsPdfAsyncToStreamWithCanceledTokenThrowsTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var range = p.Workbook.Worksheets[0].Cells["D3:F6"];
            using var ms = new MemoryStream();
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            await Assert.ThrowsExactlyAsync<TaskCanceledException>(() => range.SaveAsPdfAsync(ms, cts.Token));
        }

        [TestMethod]
        public void SaveRangeToNonWritableStreamThrowsTest()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var range = p.Workbook.Worksheets[0].Cells["D3:F6"];
            using var readOnly = new MemoryStream(new byte[16], writable: false);
            Assert.ThrowsExactly<ArgumentException>(() => range.SaveAsPdf(readOnly));
        }

        [TestMethod]
        public void ThreeFonts_NoSkip_RendersAllThreeCorrectly()
        {
            // Baseline: three different fonts, no skipping. Verifies the normal path still works
            // after the subsetting rewrite. Open the PDF and confirm A1/B1/C1 read correctly.
            using var p = OpenPackage("ThreeFonts_NoSkip.xlsx", true);
            var ws = p.Workbook.Worksheets.Add("Sheet1");

            ws.Cells["A1"].Style.Font.Name = "Aptos Narrow";
            ws.Cells["A1"].Value = "A1";
            ws.Cells["B1"].Style.Font.Name = "Times New Roman";
            ws.Cells["B1"].Value = "B1";
            ws.Cells["C1"].Style.Font.Name = "Arial";
            ws.Cells["C1"].Value = "C1";

            SaveAsPdf(ws, "ThreeFonts_NoSkip.pdf");
        }

        [TestMethod]
        public void MultiSheetWorkbook()
        {
            // Baseline: three different fonts, no skipping. Verifies the normal path still works
            // after the subsetting rewrite. Open the PDF and confirm A1/B1/C1 read correctly.
            using var p = OpenPackage("MultiSheetWorkbook.xlsx", true);
            p.Workbook.ConfigureFonts(x => x.SearchSystemDirectories = true);
            var ws = p.Workbook.Worksheets.Add("Sheet1");

            ws.Cells["A1"].Value = "Sheet1:A1";

            var ws2 = p.Workbook.Worksheets.Add("Sheet2");

            ws2.Cells["A1"].Style.Font.Name = "Times New Roman";
            ws2.Cells["A1"].Value = "Sheet2:A1";

            SaveAsPdf(p.Workbook, "MultiSheetWorkbook.pdf");
        }

        [TestMethod]
        public void MultiRanges()
        {
            // Baseline: three different fonts, no skipping. Verifies the normal path still works
            // after the subsetting rewrite. Open the PDF and confirm A1/B1/C1 read correctly.
            using var p = OpenPackage("MultiRanges.xlsx", true);
            p.Workbook.ConfigureFonts(x => x.SearchSystemDirectories = true);
            var ws = p.Workbook.Worksheets.Add("Sheet1");

            ws.Cells["A1"].Value = "Sheet1:A1";
            ws.Cells["F100"].Value = "Sheet1:F100";

            SaveAsPdf(p.Workbook, "MultiRanges.pdf", ws.Cells["A1"], ws.Cells["F100"]);
        }

        [TestMethod]
        public void SingleRange()
        {
            // Baseline: three different fonts, no skipping. Verifies the normal path still works
            // after the subsetting rewrite. Open the PDF and confirm A1/B1/C1 read correctly.
            using var p = OpenPackage("SingleRange.xlsx", true);
            p.Workbook.ConfigureFonts(x => x.SearchSystemDirectories = true);
            var ws = p.Workbook.Worksheets.Add("Sheet1");

            ws.Cells["A1"].Value = "Sheet1:A1";

            SaveAsPdf(p.Workbook, "SingleRange.pdf", ws.Cells["A1"]);
        }

        [TestMethod]
        public void ArialBlack_RendersCorrectly()
        {
            // Baseline: three different fonts, no skipping. Verifies the normal path still works
            // after the subsetting rewrite. Open the PDF and confirm A1/B1/C1 read correctly.
            using var p = OpenPackage("ArialBlack.xlsx", true);
            var ws = p.Workbook.Worksheets.Add("Sheet1");

            ws.Cells["A1"].Style.Font.Name = "Arial Black";
            ws.Cells["A1"].Value = "A1";

            SaveAsPdf(ws, "ArialBlack.pdf");
        }

        [TestMethod]
        public void ThreeFonts_SkipAll_CollapseToSharedLastResort()
        {
            // The regression case: three fonts, all skipped via OnFontEmbedding. Expected AFTER the fix:
            //   - small PDF (one shared Archivo subset, not three whole fonts)
            //   - A1 / B1 / C1 render DISTINCTLY and correctly (not all "A1")
            //   - the PDF opens without corruption
            using var p = OpenPackage("ThreeFonts_SkipAll.xlsx", true);
            var ws = p.Workbook.Worksheets.Add("Sheet1");

            ws.Cells["A1"].Style.Font.Name = "Aptos Narrow";
            ws.Cells["A1"].Value = "A1";
            ws.Cells["B1"].Style.Font.Name = "Times New Roman";
            ws.Cells["B1"].Value = "B1";
            ws.Cells["C1"].Style.Font.Name = "Arial";
            ws.Cells["C1"].Value = "C1";

            p.Workbook.ConfigureFonts(cfg =>
            {
                cfg.OnFontEmbedding(info =>
                {
                    System.Diagnostics.Debug.WriteLine("OnFontEmbedding fired for: " + info.FontName);
                    return FontEmbeddingDecision.Skip;
                });
            });


            SaveAsPdf(ws, "ThreeFonts_SkipAll.pdf");
        }

        [TestMethod]
        // works as expected.
        //[DataRow("PDFTest.xlsx", "C:\\epplustest\\pdf\\FullPageTest56.pdf", "Sheet1")]
        [DataRow("Aico_0105_S_ALR_87011990_AICO_ASSET_ITE_2025-04_BS.xlsx", "C:\\epplustest\\pdf\\OutputTest1.1.pdf", "SAP Data")]

        // Output file: OutputTest1.2.pdf
        // 1. Minus signs alignment in cells differs from Excel. ------------------------------------------------ Comment: Currently no support for number formats. Requires implementing number formats.
        // 2. Dimension seems to differ from Excel, Excel stops at row 75, EPPlus goes to row 89. --------------- Fixed
        // 3. Row headings are sligthly wider in EPPlus than in Excel. ------------------------------------------ Fixed
        [DataRow("Aico_0105_S_ALR_87011990_AICO_ASSET_ITE_2025-04_BS.xlsx", "C:\\epplustest\\pdf\\OutputTest1.2.pdf", "Summary")]
        // works as expected
        [DataRow("Aico WiP 120180 FBL3N for 0110 in 2025-04.xlsx", "C:\\epplustest\\pdf\\OutputTest1.4.pdf", "Technical")]
        [DataRow("Aico KKS1 Variance Calculation for 0105 in 2025-04 (25_4_2025 15_43_40) .xlsx", "C:\\epplustest\\pdf\\OutputTest1.5.pdf", "Technical")]

        // Output file: OutputTest1.6.pdf
        // 1. Merged cells not working ------------------------------------ Fixed. Comment Merged cells was fine, it was borders being rendered inside merged cells.
        // 2. Pattern fills looks differnt, in some cases not working -----
        // 3. Rotation of text in cells not working (the dates). ----------
        // [DataRow("R05.xlsx", "C:\\epplustest\\pdf\\OutputTest1.6.pdf", "R05 Arbeitseinteilung")]
        [DataRow("R05 - Copy.xlsx", "C:\\epplustest\\pdf\\OutputTest1.6.pdf", "R05 Arbeitseinteilung")]
        //[DataRow("PatternStyles.xlsx", "C:\\epplustest\\pdf\\OutputTest1.8.pdf", "Sheet1")]
        public void WorkbookTests(string sourceFile, string outputPath, string wsName)
        {
            using var p = OpenTemplatePackage(sourceFile);
            var ws = p.Workbook.Worksheets[wsName];
            var d = ws.Dimension;
            var d2 = ws.DimensionByValue;

            PdfPageSettings pageSettings = new PdfPageSettings(ws.Workbook.RenderContext.FontEngine);
            pageSettings.CommentsAndNotes = CommentsAndNotes.AtEndOfSheet;

            pageSettings.CellErrors = CellErrors.Displayed;
            pageSettings.Debug = true;
            pageSettings.PrintAsText = true;
            pageSettings.ShowGridLines = false;
            pageSettings.ShowHeadings = false;

            var pdfCatalog = new PdfCatalog(pageSettings, ws);
            pdfCatalog.Save(outputPath);

        }

        [TestMethod]
        public void TableDiff()
        {
            using var p = OpenTemplatePackage("TableDiff.xlsx");
            var wb = p.Workbook;
            var ws0 = wb.Worksheets[0];
            string path = _pdfPath + "TableDiff.pdf";
            wb.SaveAsPdf(path, ws0);
        }

        [TestMethod]
        public void PictureOutside()
        {
            using var p = OpenTemplatePackage("Pdf_picture_outside.xlsx");
            var wb = p.Workbook;
            var ws0 = wb.Worksheets[0];
            ws0.PrinterSettings.ShowGridLines = true;
            string path = _pdfPath + "PictureOutside.pdf";
            wb.SaveAsPdf(path, ws0);
        }

        [TestMethod]
        public void EPPlusToPdf()
        {
            string[][] pixels =
            {
                new[] { "#805840", "#805840", "#805840", "#C0A070", "#C0A070", "#C0A070", "#402820", "#402820", "#402820", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#402820", "#402820", "#805840", "#402820", "#000000", "#000000", "#805840", "#402820", "#402820", "#000000", "#402820", "#402820" },
                new[] { "#805840", "#805840", "#C0A070", "#C0A070", "#402820", "#805840", "#805840", "#402820", "#402820", "#C0A070", "#C0A070", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#402820", "#805840", "#402820", "#000000", "#805840", "#402820", "#000000", "#402820", "#402820", "#402820" },
                new[] { "#805840", "#805840", "#C0A070", "#C0A070", "#C0A070", "#402820", "#A87850", "#A87850", "#402820", "#402820", "#C0A070", "#C0A070", "#402820", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#402820", "#805840", "#402820", "#402820", "#402820", "#000000", "#402820", "#402820", "#402820", "#402820" },
                new[] { "#E0C8A0", "#E0C8A0", "#A87850", "#C0A070", "#C0A070", "#402820", "#402820", "#C0A070", "#C0A070", "#C0A070", "#C0A070", "#C0A070", "#402820", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#805840", "#805840", "#805840", "#000000", "#402820", "#402820", "#402820", "#000000" },
                new[] { "#E0C8A0", "#E0C8A0", "#E0C8A0", "#A87850", "#C0A070", "#805840", "#402820", "#402820", "#C0A070", "#A87850", "#805840", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#402820", "#402820", "#402820", "#402820", "#805840", "#805840", "#000000", "#406850", "#000000", "#000000", "#384038" },
                new[] { "#402820", "#402820", "#E0C8A0", "#E0C8A0", "#A87850", "#805840", "#402820", "#402820", "#A87850", "#C0A070", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#402820", "#402820", "#402820", "#402820", "#402820", "#406850", "#406850", "#406850", "#406850", "#000000", "#384038", "#384038" },
                new[] { "#70A070", "#70A070", "#000000", "#000000", "#E0C8A0", "#805840", "#805840", "#402820", "#A87850", "#C0A070", "#C0A070", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#402820", "#402820", "#402820", "#384038", "#384038", "#384038", "#406850", "#406850", "#406850", "#406850", "#000000", "#384038", "#384038" },
                new[] { "#70A070", "#70A070", "#70A070", "#70A070", "#406850", "#406850", "#406850", "#000000", "#000000", "#000000", "#000000", "#402820", "#402820", "#000000", "#000000", "#000000", "#000000", "#000000", "#384038", "#384038", "#384038", "#384038", "#384038", "#384038", "#384038", "#406850", "#406850", "#406850", "#000000", "#384038", "#384038", "#384038" },
                new[] { "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#384038", "#384038", "#406850", "#406850", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#384038", "#384038", "#384038", "#384038", "#384038", "#384038", "#384038", "#406850", "#406850", "#406850", "#000000", "#384038", "#384038", "#384038" },
                new[] { "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#384038", "#406850", "#406850", "#406850", "#406850", "#406850", "#406850", "#000000", "#000000", "#000000", "#384038", "#384038", "#384038", "#384038", "#384038", "#384038", "#384038", "#406850", "#406850", "#384038", "#000000", "#384038", "#384038", "#384038" },
                new[] { "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#406850", "#406850", "#406850", "#406850", "#406850", "#406850", "#406850", "#000000", "#000000", "#000000", "#384038", "#384038", "#384038", "#384038", "#384038", "#000000", "#000000", "#384038", "#384038", "#384038", "#000000", "#406850", "#384038", "#384038" },
                new[] { "#384038", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#406850", "#406850", "#406850", "#406850", "#406850", "#406850", "#406850", "#000000", "#000000", "#000000", "#384038", "#384038", "#384038", "#384038", "#384038", "#000000", "#000000", "#406850", "#406850", "#406850", "#406850", "#406850", "#384038", "#384038" },
                new[] { "#70A070", "#384038", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#406850", "#406850", "#406850", "#406850", "#406850", "#406850", "#406850", "#000000", "#000000", "#000000", "#384038", "#384038", "#384038", "#384038", "#000000", "#000000", "#406850", "#406850", "#406850", "#406850", "#406850", "#384038", "#384038" },
                new[] { "#70A070", "#70A070", "#384038", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#406850", "#406850", "#406850", "#406850", "#406850", "#406850", "#406850", "#406850", "#000000", "#000000", "#384038", "#384038", "#384038", "#384038", "#000000", "#384038", "#406850", "#406850", "#406850", "#406850", "#384038", "#384038", "#384038" },
                new[] { "#406850", "#406850", "#406850", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#384038", "#384038", "#406850", "#406850", "#406850", "#406850", "#406850", "#406850", "#000000", "#000000", "#384038", "#384038", "#384038", "#000000", "#384038", "#406850", "#406850", "#406850", "#406850", "#384038", "#384038", "#384038", "#384038" },
                new[] { "#000000", "#406850", "#406850", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#384038", "#384038", "#406850", "#406850", "#406850", "#406850", "#406850", "#000000", "#000000", "#384038", "#384038", "#000000", "#000000", "#384038", "#406850", "#406850", "#406850", "#406850", "#406850", "#384038", "#384038", "#000000", "#000000" },
                new[] { "#000000", "#406850", "#406850", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#70A070", "#384038", "#384038", "#384038", "#384038", "#406850", "#406850", "#000000", "#000000", "#384038", "#000000", "#000000", "#000000", "#384038", "#406850", "#406850", "#406850", "#406850", "#406850", "#406850", "#000000", "#000000", "#000000", "#000000" },
                new[] { "#000000", "#000000", "#000000", "#000000", "#406850", "#406850", "#406850", "#70A070", "#70A070", "#70A070", "#70A070", "#406850", "#000000", "#406850", "#406850", "#406850", "#000000", "#384038", "#000000", "#406850", "#406850", "#406850", "#406850", "#384038", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000" },
                new[] { "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#70A070", "#70A070", "#406850", "#000000", "#406850", "#406850", "#406850", "#000000", "#384038", "#000000", "#406850", "#406850", "#406850", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000" },
                new[] { "#000000", "#000000", "#402820", "#402820", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#406850", "#406850", "#406850", "#406850", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#402820", "#402820", "#805840", "#805840" },
                new[] { "#000000", "#000000", "#402820", "#402820", "#000000", "#B8A898", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#402820", "#402820", "#402820", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#805840", "#805840", "#000000", "#000000", "#000000", "#000000", "#000000", "#402820", "#805840", "#C0A070", "#805840" },
                new[] { "#000000", "#000000", "#000000", "#805840", "#805840", "#000000", "#B8A898", "#B8A898", "#000000", "#000000", "#000000", "#000000", "#402820", "#E0C8A0", "#E0C8A0", "#402820", "#402820", "#000000", "#402820", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#402820", "#402820", "#805840", "#C0A070", "#C0A070", "#805840" },
                new[] { "#000000", "#000000", "#000000", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#805840", "#000000", "#000000", "#000000", "#000000", "#000000", "#E0C8A0", "#FFF0E0", "#FFF0E0", "#402820", "#402820", "#000000", "#000000", "#402820", "#402820", "#402820", "#402820", "#402820", "#805840", "#805840", "#805840", "#805840", "#C0A070", "#C0A070", "#805840", "#805840" },
                new[] { "#000000", "#000000", "#000000", "#E0C8A0", "#FFF0E0", "#FFF0E0", "#805840", "#805840", "#805840", "#C0A070", "#C0A070", "#402820", "#E0C8A0", "#FFF0E0", "#FFF0E0", "#805840", "#805840", "#A87850", "#000000", "#000000", "#402820", "#402820", "#402820", "#402820", "#805840", "#805840", "#805840", "#C0A070", "#C0A070", "#C0A070", "#805840", "#000000" },
                new[] { "#000000", "#000000", "#000000", "#000000", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#E0C8A0", "#E0C8A0", "#C0A070", "#805840", "#E0C8A0", "#E0C8A0", "#FFF0E0", "#FFF0E0", "#805840", "#805840", "#A87850", "#805840", "#000000", "#000000", "#402820", "#805840", "#C0A070", "#C0A070", "#C0A070", "#C0A070", "#C0A070", "#C0A070", "#805840", "#000000", "#805840" },
                new[] { "#000000", "#000000", "#000000", "#000000", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#E0C8A0", "#E0C8A0", "#C0A070", "#805840", "#E0C8A0", "#E0C8A0", "#FFF0E0", "#FFF0E0", "#805840", "#A87850", "#A87850", "#A87850", "#805840", "#000000", "#402820", "#805840", "#C0A070", "#C0A070", "#C0A070", "#C0A070", "#C0A070", "#C0A070", "#805840", "#000000", "#805840" },
                new[] { "#000000", "#000000", "#000000", "#000000", "#C0A070", "#FFF0E0", "#FFF0E0", "#E0C8A0", "#E0C8A0", "#805840", "#FFF0E0", "#E0C8A0", "#E0C8A0", "#FFF0E0", "#FFF0E0", "#805840", "#A87850", "#A87850", "#A87850", "#805840", "#805840", "#000000", "#000000", "#402820", "#402820", "#C0A070", "#C0A070", "#402820", "#000000", "#000000", "#000000", "#805840" },
                new[] { "#000000", "#000000", "#000000", "#000000", "#C0A070", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#E0C8A0", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#E0C8A0", "#FFF0E0", "#FFF0E0", "#805840", "#A87850", "#A87850", "#A87850", "#A87850", "#805840", "#805840", "#402820", "#402820", "#402820", "#C0A070", "#402820", "#000000", "#000000", "#805840", "#805840", "#805840" },
                new[] { "#000000", "#000000", "#000000", "#000000", "#000000", "#C0A070", "#C0A070", "#FFF0E0", "#E0C8A0", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#E0C8A0", "#FFF0E0", "#FFF0E0", "#402820", "#A87850", "#A87850", "#A87850", "#C0A070", "#805840", "#805840", "#402820", "#402820", "#402820", "#402820", "#000000", "#000000", "#805840", "#805840", "#E0C8A0", "#A87850" },
                new[] { "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#E0C8A0", "#E0C8A0", "#E0C8A0", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#402820", "#A87850", "#A87850", "#C0A070", "#C0A070", "#805840", "#805840", "#402820", "#402820", "#000000", "#000000", "#000000", "#805840", "#805840", "#E0C8A0", "#E0C8A0", "#A87850" },
                new[] { "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#FFF0E0", "#E0C8A0", "#E0C8A0", "#FFF0E0", "#FFF0E0", "#A87850", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#402820", "#A87850", "#A87850", "#402820", "#402820", "#A87850", "#A87850", "#805840", "#805840", "#000000", "#000000", "#000000", "#805840", "#E0C8A0", "#E0C8A0", "#E0C8A0", "#A87850" },
                new[] { "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#FFF0E0", "#E0C8A0", "#E0C8A0", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#A87850", "#A87850", "#402820", "#402820", "#000000", "#000000", "#000000", "#A87850", "#A87850", "#A87850", "#A87850", "#A87850", "#000000", "#402820", "#402820", "#E0C8A0", "#C0A070", "#C0A070", "#E0C8A0", "#E0C8A0" },
                new[] { "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#FFF0E0", "#E0C8A0", "#E0C8A0", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#A87850", "#402820", "#000000", "#000000", "#000000", "#A87850", "#A87850", "#A87850", "#402820", "#402820", "#402820", "#000000", "#805840", "#805840", "#C0A070", "#C0A070", "#C0A070", "#E0C8A0", "#E0C8A0" },
                new[] { "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#A87850", "#FFF0E0", "#E0C8A0", "#FFF0E0", "#A87850", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#000000", "#000000", "#A87850", "#A87850", "#A87850", "#805840", "#000000", "#000000", "#402820", "#000000", "#805840", "#A87850", "#C0A070", "#C0A070", "#C0A070", "#C0A070", "#402820" },
                new[] { "#000000", "#000000", "#000000", "#000000", "#70A070", "#70A070", "#402820", "#FFF0E0", "#E0C8A0", "#FFF0E0", "#000000", "#805840", "#805840", "#402820", "#000000", "#000000", "#805840", "#805840", "#000000", "#000000", "#000000", "#000000", "#000000", "#402820", "#000000", "#000000", "#A87850", "#C0A070", "#C0A070", "#C0A070", "#C0A070", "#402820" },
                new[] { "#000000", "#000000", "#000000", "#70A070", "#70A070", "#70A070", "#406850", "#000000", "#E0C8A0", "#FFF0E0", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#C0A070", "#C0A070", "#C0A070", "#C0A070", "#402820", "#A87850" },
                new[] { "#000000", "#000000", "#406850", "#406850", "#406850", "#406850", "#406850", "#000000", "#E0C8A0", "#FFF0E0", "#000000", "#000000", "#000000", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#805840", "#805840", "#805840", "#805840", "#000000", "#000000", "#000000", "#000000", "#C0A070", "#E0C8A0", "#C0A070", "#402820", "#A87850", "#A87850" },
                new[] { "#000000", "#406850", "#406850", "#406850", "#406850", "#406850", "#406850", "#000000", "#402820", "#A87850", "#805840", "#000000", "#FFF0E0", "#FFF0E0", "#A87850", "#A87850", "#A87850", "#A87850", "#805840", "#805840", "#805840", "#805840", "#402820", "#402820", "#000000", "#000000", "#E0C8A0", "#E0C8A0", "#C0A070", "#402820", "#A87850", "#A87850" },
                new[] { "#406850", "#406850", "#406850", "#406850", "#406850", "#406850", "#406850", "#000000", "#000000", "#A87850", "#805840", "#E0C8A0", "#E0C8A0", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#805840", "#805840", "#805840", "#000000", "#000000", "#E0C8A0", "#E0C8A0", "#402820", "#A87850", "#A87850", "#384038" },
                new[] { "#406850", "#406850", "#406850", "#406850", "#406850", "#406850", "#406850", "#000000", "#000000", "#000000", "#805840", "#A87850", "#402820", "#402820", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#805840", "#402820", "#000000", "#000000", "#E0C8A0", "#000000", "#000000", "#A87850", "#384038", "#384038" },
                new[] { "#406850", "#406850", "#406850", "#406850", "#406850", "#406850", "#406850", "#000000", "#000000", "#000000", "#402820", "#E0C8A0", "#402820", "#402820", "#E0C8A0", "#E0C8A0", "#E0C8A0", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#C0A070", "#000000", "#000000", "#000000", "#A87850", "#000000", "#406850" },
                new[] { "#384038", "#384038", "#384038", "#384038", "#406850", "#406850", "#406850", "#000000", "#000000", "#000000", "#402820", "#A87850", "#E0C8A0", "#E0C8A0", "#E0C8A0", "#E0C8A0", "#E0C8A0", "#A87850", "#000000", "#000000", "#000000", "#000000", "#000000", "#805840", "#000000", "#C0A070", "#000000", "#000000", "#000000", "#000000", "#406850", "#406850" },
                new[] { "#70A070", "#406850", "#406850", "#384038", "#406850", "#406850", "#000000", "#000000", "#000000", "#000000", "#402820", "#A87850", "#A87850", "#E0C8A0", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#A87850", "#000000", "#000000", "#000000", "#000000", "#000000", "#805840", "#805840", "#000000", "#000000", "#000000", "#000000", "#406850", "#406850", "#406850" },
                new[] { "#70A070", "#70A070", "#406850", "#000000", "#406850", "#406850", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#A87850", "#A87850", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#E0C8A0", "#000000", "#000000", "#000000", "#000000", "#805840", "#805840", "#000000", "#000000", "#000000", "#000000", "#406850", "#406850", "#406850", "#406850" },
                new[] { "#70A070", "#70A070", "#70A070", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#A87850", "#000000", "#A87850", "#FFF0E0", "#FFF0E0", "#FFF0E0", "#E0C8A0", "#000000", "#000000", "#000000", "#000000", "#805840", "#000000", "#000000", "#402820", "#000000", "#406850", "#406850", "#406850", "#406850", "#406850" },
                new[] { "#000000", "#000000", "#70A070", "#384038", "#406850", "#406850", "#406850", "#000000", "#000000", "#000000", "#000000", "#A87850", "#000000", "#402820", "#402820", "#402820", "#402820", "#402820", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#402820", "#402820", "#000000", "#406850", "#406850", "#406850", "#000000", "#000000" },
                new[] { "#384038", "#384038", "#70A070", "#70A070", "#384038", "#406850", "#406850", "#000000", "#000000", "#000000", "#000000", "#000000", "#A87850", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#402820", "#402820", "#000000", "#406850", "#384038", "#000000", "#000000", "#000000", "#000000" },
                new[] { "#384038", "#384038", "#000000", "#70A070", "#406850", "#406850", "#406850", "#406850", "#000000", "#000000", "#000000", "#000000", "#A87850", "#A87850", "#A87850", "#A87850", "#A87850", "#000000", "#000000", "#000000", "#000000", "#000000", "#000000", "#402820", "#402820", "#000000", "#406850", "#384038", "#000000", "#000000", "#000000", "#000000" },
            };
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Snake");
            const double rowHeightPts = 15;
            double columnWidth = rowHeightPts * (96.0 / 72.0) / 7.0;
            ws.Column(1).Width = columnWidth;
            ws.Cells["A1"].Value = "SOLID";
            ws.Cells["AE50"].Value = "SNAKE";
            //ws.Cells["AF51"].Value = " ";
            const int startRow = 2;
            const int startCol = 1; // D
            for (int y = 0; y < pixels.Length; y++)
            {
                for (int x = 0; x < pixels[y].Length; x++)
                {
                    var cell = ws.Cells[startRow + y, startCol + x];
                    var color = System.Drawing.ColorTranslator.FromHtml(pixels[y][x]);
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(color);
                    ws.Column(startCol + x).Width = columnWidth;
                }
                ws.Row(startRow + y).Height = 15;
            }
            ws.PrinterSettings.TopMargin = 0.1d;
            ws.PrinterSettings.BottomMargin = 0.1d;
            ws.PrinterSettings.LeftMargin = 0.1d;
            ws.PrinterSettings.RightMargin = 0.1d;
            ws.PrinterSettings.HorizontalCentered = true;
            ws.PrinterSettings.VerticalCentered = true;
            CreatePathIfNotExists(_pdfPath);

            p.Workbook.SaveAsPdf(_pdfPath + "Snake.Pdf");
            p.SaveAs(_pdfPath + "Snake.xlsx");
        }

        [TestMethod]
        public void Testing()
        {
            using var p = OpenTemplatePackage("PDFTestKarl.xlsx");
            var wb = p.Workbook;
            string path = _pdfPath + "WorksheetTest1.pdf";
            wb.SaveAsPdf(path);
            AssertLooksLikePdf(File.ReadAllBytes(path));
        }

        [TestMethod]
        public void EachWorksheetUsesItsOwnOrientation()
        {
            using (var package = OpenTemplatePackage("PDFTestKarl.xlsx"))
            {
                package.Workbook.Worksheets[0].PrinterSettings.Orientation = eOrientation.Portrait;
                package.Workbook.Worksheets[1].PrinterSettings.Orientation = eOrientation.Landscape;

                var settings = GetPdfSettings.GetPdfSettingsFromPrinterSettings(
                    package.Workbook,
                    package.Workbook.Worksheets[0].PrinterSettings);

                byte[] pdf;
                using (var ms = new MemoryStream())
                {
                    var pdfCatalog = new PdfCatalog(settings, package.Workbook);
                    pdfCatalog.Save(ms);
                    pdf = ms.ToArray();
                }

                var matches = Regex.Matches(
                    Encoding.ASCII.GetString(pdf),
                    @"/MediaBox\s*\[\s*0\s+0\s+(?<w>[\d.]+)\s+(?<h>[\d.]+)\s*\]");

                Assert.AreEqual(2, matches.Count, "Expected one page per worksheet.");

                var ci = CultureInfo.InvariantCulture;
                double w1 = double.Parse(matches[0].Groups["w"].Value, ci);
                double h1 = double.Parse(matches[0].Groups["h"].Value, ci);
                double w2 = double.Parse(matches[1].Groups["w"].Value, ci);
                double h2 = double.Parse(matches[1].Groups["h"].Value, ci);

                Assert.IsTrue(h1 > w1, "Page 1 should be portrait.");
                Assert.IsTrue(w2 > h2, "Page 2 should be landscape.");
                // Landscape is the same paper transposed, not a different paper size.
                Assert.AreEqual(w1, h2, 0.01d);
                Assert.AreEqual(h1, w2, 0.01d);
            }
        }

        [TestMethod]
        public void EachWorksheetUsesItsOwnShowGridLines()
        {
            using (var package = OpenTemplatePackage("PDFTestKarl.xlsx"))
            {
                package.Workbook.Worksheets[0].PrinterSettings.ShowGridLines = false;
                package.Workbook.Worksheets[1].PrinterSettings.ShowGridLines = true;

                var baseSettings = GetPdfSettings.GetPdfSettingsFromPrinterSettings(
                    package.Workbook,
                    package.Workbook.Worksheets[0].PrinterSettings);

                var s0 = GetPdfSettings.GetPdfSettingsForSheet(
                    baseSettings, package.Workbook.Worksheets[0].PrinterSettings);
                var s1 = GetPdfSettings.GetPdfSettingsForSheet(
                    baseSettings, package.Workbook.Worksheets[1].PrinterSettings);

                Assert.IsFalse(s0.ShowGridLines, "Sheet 1 did not ask for gridlines.");
                Assert.IsTrue(s1.ShowGridLines, "Sheet 2 asked for gridlines.");
                Assert.IsFalse(baseSettings.ShowGridLines, "The base object must not be mutated.");
            }
        }

        [TestMethod]
        public void EachWorksheetUsesItsOwnPaperSize()
        {
            using (var package = OpenTemplatePackage("PDFTestKarl.xlsx"))
            {
                // Orientation is set explicitly so a transposed page size cannot be
                // mistaken for a different paper size.
                package.Workbook.Worksheets[0].PrinterSettings.Orientation = eOrientation.Portrait;
                package.Workbook.Worksheets[1].PrinterSettings.Orientation = eOrientation.Portrait;
                package.Workbook.Worksheets[0].PrinterSettings.PaperSize = ePaperSize.A4;
                package.Workbook.Worksheets[1].PrinterSettings.PaperSize = ePaperSize.A3;

                var settings = GetPdfSettings.GetPdfSettingsFromPrinterSettings(
                    package.Workbook,
                    package.Workbook.Worksheets[0].PrinterSettings);

                byte[] pdf;
                using (var ms = new MemoryStream())
                {
                    var pdfCatalog = new PdfCatalog(settings, package.Workbook);
                    pdfCatalog.Save(ms);
                    pdf = ms.ToArray();
                }

                var matches = Regex.Matches(
                    Encoding.ASCII.GetString(pdf),
                    @"/MediaBox\s*\[\s*0\s+0\s+(?<w>[\d.]+)\s+(?<h>[\d.]+)\s*\]");

                Assert.AreEqual(2, matches.Count, "Expected one page per worksheet.");

                var ci = CultureInfo.InvariantCulture;
                double w1 = double.Parse(matches[0].Groups["w"].Value, ci);
                double h1 = double.Parse(matches[0].Groups["h"].Value, ci);
                double w2 = double.Parse(matches[1].Groups["w"].Value, ci);
                double h2 = double.Parse(matches[1].Groups["h"].Value, ci);

                // Compare against the source of truth rather than literal point values.
                // PdfPageSize rounds mm to whole points, so 210x297 mm becomes 595x842.
                Assert.AreEqual(PdfPageSize.A4.WidthPu, w1, "Page 1 should be A4.");
                Assert.AreEqual(PdfPageSize.A4.HeightPu, h1, "Page 1 should be A4.");
                Assert.AreEqual(PdfPageSize.A3.WidthPu, w2, "Page 2 should be A3, not sheet 1's A4.");
                Assert.AreEqual(PdfPageSize.A3.HeightPu, h2, "Page 2 should be A3, not sheet 1's A4.");
            }
        }

        [TestMethod]
        public void HeaderFooterTest1()
        {
            using var p = OpenTemplatePackage("1.06-Salesreport.xlsx");
            var ws = p.Workbook.Worksheets[0];
            string path = _pdfPath + "HeaderFooterTest1.pdf";
            ws.SaveAsPdf(path);
            Assert.IsTrue(File.Exists(path), "PDF file was not created.");
            AssertLooksLikePdf(File.ReadAllBytes(path));
        }

        [TestMethod]
        public void GetOriginX_CenteringOff_ReturnsContentBoundsLeft()
        {
            var s = new PdfPageSettings(null);
            var p = new Page()
            {
                FromRow = 1,
                ToRow = 10,
                FromColumn = 1,
                ToColumn = 5,
                UsedWidth = 100,
                UsedHeight = 100,
                RowHeights = new double[10]
            };

            Assert.AreEqual(s.ContentBounds.Left, PdfLayout.GetOriginX(s, p), 0.0001);
        }

        [TestMethod]
        public void GetOrigin_FlagsAreIndependent()
        {
            var s = new PdfPageSettings(null);
            s.CenterOnPageHorizontally = true;
            var p = new Page()
            {
                FromRow = 1,
                ToRow = 10,
                FromColumn = 1,
                ToColumn = 5,
                UsedWidth = s.ContentBounds.Width - 100d,
                UsedHeight = s.ContentBounds.Height - 200d,
                RowHeights = new double[10]
            };
            Assert.AreEqual(s.ContentBounds.Left + 50d, PdfLayout.GetOriginX(s, p), 0.0001);
            Assert.AreEqual(s.ContentBounds.Top, PdfLayout.GetOriginY(s, p), 0.0001);
        }

        [TestMethod]
        public void GetClampedCellWidth_CellFitsWithinPage_ReturnsCellWidthUnchanged()
        {
            var s = new PdfPageSettings(null);

            Assert.AreEqual(51.71d, PdfLayout.GetClampedCellWidth(s, 126.31d, 51.71d), 0.0001);
        }

        [TestMethod]
        public void LargeTableTest1()
        {
            using var p = OpenTemplatePackage("BlazorSample1 (12).xlsx");
            var ws = p.Workbook.Worksheets[1];
            string path = _pdfPath + "LargeTableTest.pdf";
            ws.SaveAsPdf(path);
            Assert.IsTrue(File.Exists(path), "PDF file was not created.");
            AssertLooksLikePdf(File.ReadAllBytes(path));
        }

        [TestMethod]
        public void headerFooterImage()
        {
            using var p = OpenTemplatePackage("EPPlus Sample 3.xlsx");
            var ws = p.Workbook.Worksheets[0];
            string path = _pdfPath + "EPPlus Sample 3.pdf";
            ws.SaveAsPdf(path);
        }

        [TestMethod]
        public void headerFooterImage2()
        {
            using var p = CreateWorkbook();
            var ws = p.Workbook.Worksheets[0];
            string path = _pdfPath + "EPPlus Sample 3.pdf";
            ws.SaveAsPdf(path);
        }
        public ExcelPackage CreateWorkbook()
        {
            InitDataTable();
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Html export sample 2");
            var tableRange = sheet.Cells["A1"].LoadFromDataTable(_dataTable, true, TableStyles.Dark3);

            //Configure the table
            var table = sheet.Tables.GetFromRange(tableRange);
            table.Sort(x => x.SortBy.ColumnNamed("Population", eSortOrder.Descending));
            table.ShowTotal = true;
            table.Columns[0].TotalsRowLabel = "Total";
            table.Columns[1].TotalsRowFunction = RowFunctions.Sum;
            table.Columns[2].TotalsRowFunction = RowFunctions.Sum;

            //Add column for population density
            table.Columns.Add(1);
            tableRange = table.Range;
            table.Columns[3].CalculatedColumnFormula = $"{table.Name}[[#This Row],[Population]]/{table.Name}[[#This Row],[Area (km2)]]";
            table.Columns[3].Name = "Density";
            table.Columns[3].TotalsRowFunction = RowFunctions.Average;
            sheet.Calculate();

            //// format the header
            table.Range.TakeColumnsBetween(1, 3).SkipRows(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

            // format the rows
            var lastDataRow = tableRange.End.Row - 1;
            sheet.Cells[tableRange.Start.Row, 1, lastDataRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            sheet.Cells[tableRange.Start.Row, 2, lastDataRow, 2].Style.Numberformat.Format = "#,##0";
            sheet.Cells[tableRange.Start.Row, 3, lastDataRow, 3].Style.Numberformat.Format = "#,##0 \"km2\"";
            sheet.Cells[tableRange.Start.Row, 4, lastDataRow, 4].Style.Numberformat.Format = "#,##0.0";

            // format the total row
            var totalRow = tableRange.End.Row;
            sheet.Cells[totalRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            sheet.Cells[totalRow, 2].Style.Numberformat.Format = "#,##0";
            sheet.Cells[totalRow, 3].Style.Numberformat.Format = "#,##0 \"km2\"";
            sheet.Cells[totalRow, 4].Style.Numberformat.Format = "\"Avg: \"#,##0.0 ";
            sheet.Cells.AutoFitColumns();

            //Set the header and footer values
            var text = sheet.HeaderFooter.OddHeader.Centered.AddText("EPPlus Sample 3");
            text.FontSize = 18;
            var imageFile = Path.Combine(_imagePath, "EPPlus-logo-small.jpg");
            if (File.Exists(imageFile))
            {
                //sheet.HeaderFooter.OddHeader.LeftAligned.AddText("Logo:");
                sheet.HeaderFooter.OddHeader.LeftAligned.AddImage(new FileInfo(imageFile));
            }

            sheet.HeaderFooter.OddFooter.Centered.AddPageNumber();
            sheet.HeaderFooter.OddFooter.Centered.AddText(" of ");
            sheet.HeaderFooter.OddFooter.Centered.AddNumberOfPages();
            return package;
        }
        private DataTable _dataTable = null;

        private void InitDataTable()
        {
            if (_dataTable != null) return;
            _dataTable = new DataTable();
            _dataTable.Columns.Add("Country", typeof(string));
            _dataTable.Columns.Add("Population", typeof(int));
            var areaCol = _dataTable.Columns.Add("Area", typeof(int));
            areaCol.Caption = "Area (km2)";


            _dataTable.Rows.Add("Sweden", 10409248, 450295);
            _dataTable.Rows.Add("Norway", 5402171, 385178);
            _dataTable.Rows.Add("Netherlands", 17553530, 41198);
            _dataTable.Rows.Add("Finland", 5541806, 338145);
            _dataTable.Rows.Add("Belgium", 11521238, 30510);
            _dataTable.Rows.Add("Denmark", 5850189, 44493);
            _dataTable.Rows.Add("Lithuania", 2801264, 65300);
            _dataTable.Rows.Add("Greece", 10718565, 131940);
            _dataTable.Rows.Add("Russia", 145734038, 3972400);
            _dataTable.Rows.Add("Germany", 83124418, 357386);
            _dataTable.Rows.Add("France", 64990511, 551695);
            _dataTable.Rows.Add("Czech Republic", 10665677, 78866);
            _dataTable.Rows.Add("Slovakia", 5459781, 49036);
            _dataTable.Rows.Add("Spain", 47394223, 498468);
            _dataTable.Rows.Add("Portugal", 10256193, 91568);
            _dataTable.Rows.Add("United Kingdom", 67141684, 242495);
            _dataTable.Rows.Add("Poland", 37921592, 312685);
            _dataTable.Rows.Add("Albania", 2882740, 28748);
            _dataTable.Rows.Add("Estonia", 1322920, 45339);
            _dataTable.Rows.Add("Hungary", 9707499, 93030);
            _dataTable.Rows.Add("Romania", 19186000, 238397);
            _dataTable.Rows.Add("Italy", 60627291, 301338);
            _dataTable.Rows.Add("Bulgaria", 7051608, 110994);
            _dataTable.Rows.Add("Belarus", 9452617, 207600);
            _dataTable.Rows.Add("Austria", 8891388, 83858);
            _dataTable.Rows.Add("Switzerland", 8525611, 41290);
            _dataTable.Rows.Add("Ireland", 4818690, 70273);
            _dataTable.Rows.Add("Ukraine", 44246156, 603628);
            _dataTable.Rows.Add("Iceland", 336713, 102775);
            _dataTable.Rows.Add("Serbia", 6871547, 77453);
            _dataTable.Rows.Add("Croatia", 4156405, 56594);
            _dataTable.Rows.Add("Latvia", 1928459, 64589);
            _dataTable.Rows.Add("Bosnia and Herzegovina", 3323925, 51129);
            _dataTable.Rows.Add("Montenegro", 627809, 13812);
            _dataTable.Rows.Add("Cyrprus", 1189265, 9251);
            _dataTable.Rows.Add("Kosovo", 1798506, 10908);
            _dataTable.Rows.Add("Slovenia", 2055496, 20273);
            _dataTable.Rows.Add("Moldova", 4033963, 33846);
            _dataTable.Rows.Add("North Macedonia", 2083374, 25713);
            _dataTable.Rows.Add("United States", 331002651, 9833517);
            _dataTable.Rows.Add("China", 1412600000, 9596961);
            _dataTable.Rows.Add("India", 1417173173, 3287263);
            _dataTable.Rows.Add("Japan", 125681593, 377975);
            _dataTable.Rows.Add("Brazil", 214326223, 8515767);
            _dataTable.Rows.Add("Canada", 38246108, 9984670);
            _dataTable.Rows.Add("Australia", 25788215, 7692024);
            _dataTable.Rows.Add("Mexico", 126014024, 1964375);
            _dataTable.Rows.Add("South Africa", 59308690, 1221037);
            _dataTable.Rows.Add("Egypt", 104258327, 1001450);
            _dataTable.Rows.Add("Nigeria", 218541212, 923768);
            _dataTable.Rows.Add("Argentina", 45808747, 2780400);
            _dataTable.Rows.Add("Indonesia", 273523615, 1904569);
            _dataTable.Rows.Add("South Korea", 51780579, 100210);
            _dataTable.Rows.Add("Turkey", 84339067, 783562);
            _dataTable.Rows.Add("Saudi Arabia", 34813871, 2149690);
            _dataTable.Rows.Add("Israel", 9053300, 20770);
            _dataTable.Rows.Add("New Zealand", 5084300, 268838);
            _dataTable.Rows.Add("Thailand", 69950850, 513120);
            _dataTable.Rows.Add("Vietnam", 98168833, 331212);
        }
    }
}
