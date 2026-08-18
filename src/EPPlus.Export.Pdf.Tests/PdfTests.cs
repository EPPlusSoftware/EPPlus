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
using OfficeOpenXml;
using OfficeOpenXml.Export.PdfExport;
using OfficeOpenXml.Export.PdfExport.Settings;
using OfficeOpenXml.Style;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace EPPlusTest.PDF
{
    [TestClass]
    public class PdfTests : TestBase
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

        protected static string pdfPath = _worksheetPath + "\\PDF\\";

        [TestMethod]
        public void SaveWorksheetAsPdfTest1()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var ws = p.Workbook.Worksheets[0];
            string path = pdfPath + "WorksheetTest1.pdf";
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
            string path = pdfPath + "WorksheetTest2.pdf";
            ws.SaveAsPdf(path);
            Assert.IsTrue(File.Exists(path), "PDF file was not created.");
            AssertLooksLikePdf(File.ReadAllBytes(path));
        }

        [TestMethod]
        public void SaveRangeAsPdfTest1()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var range = p.Workbook.Worksheets[0].Cells["D3:F6"];
            string path = pdfPath + "RangeTest1.pdf";
            range.SaveAsPdf(path);
            Assert.IsTrue(File.Exists(path), "PDF file was not created.");
            AssertLooksLikePdf(File.ReadAllBytes(path));
        }

        [TestMethod]
        public void SaveWorkbookAsPdfTest1()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            string path = pdfPath + "WorkbookTest1.pdf";
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
            string path = pdfPath + "WorksheetsTest2.pdf";
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
            string path = pdfPath + "WorksheetsTest1.pdf";
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
            string path = pdfPath + "RangesTest1.pdf";
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
            _ = new PdfCatalog(ms, pageSettings, ws);
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

            PdfCatalog catalog = new PdfCatalog(outputPath, pageSettings, ws);

        }

        [TestMethod]
        public void TableDiff()
        {
            using var p = OpenTemplatePackage("TableDiff.xlsx");
            var wb = p.Workbook;
            var ws0 = wb.Worksheets[0];
            string path = pdfPath + "TableDiff.pdf";
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
            p.Workbook.SaveAsPdf(pdfPath + "Snake.Pdf");
            p.SaveAs(pdfPath + "Snake.xlsx");
        }

        [TestMethod]
        public void Testing()
        {
            using var p = OpenTemplatePackage("PDFTestKarl.xlsx");
            var wb = p.Workbook;
            string path = pdfPath + "WorksheetTest1.pdf";
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
                    new PdfCatalog(ms, settings, package.Workbook);
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

                var settings = GetPdfSettings.GetPdfSettingsFromPrinterSettings(
                    package.Workbook,
                    package.Workbook.Worksheets[0].PrinterSettings);

                byte[] pdf;
                using (var ms = new MemoryStream())
                {
                    new PdfCatalog(ms, settings, package.Workbook);
                    pdf = ms.ToArray();
                }

                var text = Encoding.ASCII.GetString(pdf);

                // One page per worksheet - guards against the count assertion below
                // being thrown off by pagination or a comments page.
                Assert.AreEqual(2, Regex.Matches(text, @"/MediaBox\s*\[").Count,
                    "Expected one page per worksheet.");

                // PdfContentStream.AddInnerGridLines writes this marker whenever the flag
                // is set, so the number of occurrences equals the number of pages that
                // asked for gridlines. Exactly one means the flag was read per sheet:
                // reading sheet 1's flag globally    would give 0, applying it to every page
                // would give 2.
                Assert.AreEqual(1, Regex.Matches(text, @"% Gridlines Start").Count,
                    "Only sheet 2 asked for gridlines.");

                //var text = Encoding.ASCII.GetString(pdf);
                //Debug.WriteLine($"pages={Regex.Matches(text, @"/MediaBox\s*\[").Count} " +
                //                $"gridlines={Regex.Matches(text, @"% Gridlines Start").Count}");
            }
        }
    }
}
