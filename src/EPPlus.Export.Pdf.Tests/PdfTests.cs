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
using OfficeOpenXml.Export.PdfExport;
using EPPlus.Export.Pdf.Settings.PdfPageSizes;
using EPPlus.Export.Pdf;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.RichText;
using OfficeOpenXml.Style;
using System.Diagnostics;

namespace EPPlusTest.PDF
{
    [TestClass]
    public class PdfTests : TestBase
    {
        /* BIG PDF TODO
         * 
         * Missing Features:
         * Embedding pictures as drawings
         * Embedding pictures as cell content
         * Embedding pictures as header/footer
         * Number formatting
         * Scaling document
         * Conditional Formatting icons
         * Vertical text
         * Equations
         * Pivot tables
         * Shapes
         * Charts
         * 3D models
         * Compression
         * Unique pdf settings per worksheet
         * 
         * Bugs:
         * Merged cell uses full width/height even when columns/rows are hidden.
         * Conditional formatting not worksing 100% of the time.
         * Table could sometimes select wrong style
         * 
         * Other
         * Cells: Remove stroke from solid fill and adjust size and position of cell more precise and only use fill command.
         * Gradients: Make diamond gradients instead of radial gradtient in from corner and center gradient
         * Text: Set up to use back up techniques when not embedding font
         * Borders: Adjust and make border look better
         * 
         */

        protected static string pdfPath = _worksheetPath + "\\PDF\\";

        //ta bort
        [TestMethod]
        public void ReadPrintAreas()
        {
            //using var p = OpenTemplatePackage("PdfPrintAreas.xlsx");
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            //using var p = OpenTemplatePackage("PDFTest_old.xlsx");
            //using var p = OpenTemplatePackage("DoubleBorder.xlsx");
            var ws = p.Workbook.Worksheets[0];
            PdfPageSettings pageSettings = new PdfPageSettings();
            pageSettings.CommentsAndNotes = CommentsAndNotes.AtEndOfSheet;
            pageSettings.CellErrors = CellErrors.NA;
            pageSettings.Debug = true;
            pageSettings.PrintAsText = true;
            pageSettings.ShowGridLines = true;
            pageSettings.ShowHeadings = true;
            PdfCatalog catlog = new PdfCatalog("C:\\epplustest\\pdf\\FullPageTest59.pdf", pageSettings, ws);
        }

        [TestMethod]
        public void SaveWorksheetAsPdfTest1()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var ws = p.Workbook.Worksheets[0];
            ws.SaveAsPdf(pdfPath + "WorksheetTest1.pdf");
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
            ws.SaveAsPdf(pdfPath + "WorksheetTest2.pdf");
        }

        [TestMethod]
        public void SaveRangeAsPdfTest1()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var range = p.Workbook.Worksheets[0].Cells["D3:F6"];
            range.SaveAsPdf(pdfPath + "RangeTest1.pdf");
        }

        [TestMethod]
        public void SaveWorkbookAsPdfTest1()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            wb.SaveAsPdf(pdfPath + "WorkbookTest1.pdf");
        }

        [TestMethod]
        public void SaveWorksheetsAsPdfTest2()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            var ws0 = wb.Worksheets[0];
            var ws1 = wb.Worksheets[1];
            var ws2 = wb.Worksheets[2];
            wb.SaveAsPdf(pdfPath + "WorksheetsTest2.pdf", ws0, ws1, ws2);
        }

        [TestMethod]
        public void SaveWorksheetsAsPdfTest1()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var wb = p.Workbook;
            var ws0 = wb.Worksheets[0];
            var ws2 = wb.Worksheets[2];
            wb.SaveAsPdf(pdfPath + "WorksheetsTest1.pdf", ws0, ws2);
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
            wb.SaveAsPdf(pdfPath + "RangesTest1.pdf", r1, r2, r3, r4);
        }

        [TestMethod]
        public void PerformanceTest()
        {
            var sw = new Stopwatch();
            Console.WriteLine("Starting...");
            sw.Start();
            using var p = OpenTemplatePackage("Aico_0105_S_ALR_87011990_AICO_ASSET_ITE_2025-04_BS.xlsx");
            Console.WriteLine($"Read package time elapsed: {sw.ElapsedMilliseconds} ms");
            sw.Restart();
            var ws = p.Workbook.Worksheets[0];
            var dun = ws.Dimension;
            var pageSettings = new PdfPageSettings
            {
                CommentsAndNotes = CommentsAndNotes.AtEndOfSheet,
                CellErrors = CellErrors.NA,
                Debug = true,
                PrintAsText = true,
                ShowGridLines = false,
                ShowHeadings = true
            };
            PdfCatalog catalog = new PdfCatalog("C:\\epplustest\\pdf\\OutputTest1.3.pdf", pageSettings, ws);
            Console.WriteLine($"workbook exported time elapsed: {sw.ElapsedMilliseconds} ms");
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

            PdfPageSettings pageSettings = new PdfPageSettings();
            pageSettings.CommentsAndNotes = CommentsAndNotes.AtEndOfSheet;

            pageSettings.CellErrors = CellErrors.Displayed;
            pageSettings.Debug = true;
            pageSettings.PrintAsText = true;
            pageSettings.ShowGridLines = false;
            pageSettings.ShowHeadings = false;

            PdfCatalog catalog = new PdfCatalog(outputPath, pageSettings, ws);
        }
    }
}
