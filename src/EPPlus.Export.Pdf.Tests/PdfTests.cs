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
         * This should be turned into tickets on github..
         * 
         * FEATURES:
         * Cells: Icons
         * Cells: Pictures in cells : Prio 3
         * Cells: Conditional Formatting
         * Cells: Calculate clip rect for merged cells with content bounds in mind
         * 
         * Text: Wrap text
         * Text: Vertical text again
         * Text: Equations : Prio 5
         * Text: Shrink To Fit (Scaling)
         * 
         * Layout: Scaling
         * Layout: Fit to number of pages (Scaling)
         * Layout: Excel Workbook to pdf (Export)
         * Layout: Selected worksheets to pdf (Export)
         * Layout: Cell range to pdf (Export)
         * Layout: Print Area (This is like a collection of cell ranges) (Export)
         * Layout: Print titles
         * Layout: Background (Excel does not add this to pdf it seems) (Images) : Prio 4
         * Layout: Remove empty last page
         * Layout: Autofit row
         * 
         * Pivot Table: Pivot table implementation : Prio 5
         * 
         * Drawings: Pictures (Images) :  Prio 3
         * Drawings Shapes (Images) : Prio 5
         * Drawings: Charts (Images) : Prio 5
         * Drawings: 3D models (Images) : Prio 5
         * 
         * PDF: Compress parts
         * 
         * IMPROVEMENTS;
         * Cells: Remove stroke from solid fill and adjust size and position of cell more precise and only use fill command.
         * 
         * Tables: Table Implementation
         * Tables: Fix order of applying table styling. Right now Table is used if it exsists, but cell styling should be prioritised and table should be ignored if cell has styluing.
         * 
         * Patterns: Adjust and make patterns look better. Fixed: DarkUp, DarkDown
         * 
         * Gradients: Make diamond gradients instead of radial gradtient in from corner and center gradient
         * 
         * Text: Set up to use back up techniques when not embedding font
         * 
         * Layout: Calculate height of cells more correctly
         *
         * Borders: Adjust and make border look better
         */

        /* REFACTOR:
         * 1. Collect: Gather all text in all worksheets, headerfooter and comments
         * 2. Shape  : Shape the Text in all worksheets, headerfooter and comments
         * 3. Layout : Autofit row, add print titles, add row and column headings, set up print areas, page breaks
         * 4. Pages: : Create number of pages needed, assign elements to pages
         * 5. PDF    : Export to pdf
         * 
         * 
         * --------------0. Set up ---------------------
         * Create class catalog that takes in a workbook, or a collection of worksheets, or a range.
         * When step1 is done we check in a comments collections if there are comments and process the comments worksheet the same way.
         * for each worksheet do following:
         * --------------1. Collect---------------------
         * Create array map of worksheet with coords, text, fills and so on.
         * This is done when all worksheets have been mapped.
         * Comments and notes will create a temporary worksheet in current workbook with all notes and comments inside cells.
         * --------------2. Shape  ---------------------
         * Go through all text in all workbbooks and shape the text.
         * We also check text height measurements and adjust row height.
         * --------------3. Layout ---------------------
         * Precalculate content area of each page using content area and page breaks and print areas. We need to take scaling into account here!
         * Add pictures, shapes, notes and other elements that lies on top of cells.
         * Then loop the collection and insert row and column headings, print titles.
         * We also need to create gridlines here
         * --------------4. Pages ----------------------
         * Create Pages and create transform objects of each element from the array and assign them to their respective page.
         * --------------5. PDF ------------------------
         * When all worksheets have been processed, combine all into one catalog and create the pdf document.
         */

        /// <summary>
        /// Old Test
        /// </summary>
        //[TestMethod]
        //public void TestWritePdf()
        //{
        //    using var p = OpenTemplatePackage("PDFTest.xlsx");
        //    //using var p = OpenTemplatePackage("PdfGrids\\PdfTextTest.xlsx");
        //    //using var p = OpenTemplatePackage("PdfGrids\\PdfPageBreakTest.xlsx");
        //    //using var p = OpenTemplatePackage("PDFTest - Copy (2).xlsx");
        //    //using var p = OpenTemplatePackage("PdfBorders.xlsx");
        //    //using var p = OpenTemplatePackage("PdfGrids\\3 2 Page Crazy Cells.xlsx");
        //    //using var p = OpenTemplatePackage("PdfGrids\\3 2 Page Crazy Cells Merged.xlsx");
        //    //using var p = OpenTemplatePackage("Gradient.xlsx");
        //    //using var p = OpenTemplatePackage("PatternFill.xlsx");
        //    var ws = p.Workbook.Worksheets[0];
        //    //var ws = p.Workbook.Worksheets[1];
        //    PdfPageSettings pageSettings = new PdfPageSettings();
        //    pageSettings.ShowGridLines = true;
        //    pageSettings.PageSize = PdfPageSize.A4;
        //    pageSettings.Orientation = Orientations.Portrait;
        //    pageSettings.Margins = PdfMargins.Normal;
        //    pageSettings.ShowGridLines = true;
        //    pageSettings.CenterOnPageHorizontally = true;
        //    pageSettings.CenterOnPageVertically = true;
        //    pageSettings.ShowHeadings = true;
        //    pageSettings.CommentsAndNotes = CommentsAndNotes.AtEndOfSheet;
        //    //Debug Flags
        //    pageSettings.Debug = true;
        //    pageSettings.PrintAsText = true;

        //    ExcelPdf pedeef = new ExcelPdf(ws, pageSettings);
        //    pedeef.CreatePdf("c:\\epplustest\\pdf\\FullPageTest49.pdf");
        //}

        /// <summary>
        /// Old Test
        /// </summary>
        //[TestMethod]
        //public void TestWritePdf2()
        //{
        //    using var p = OpenTemplatePackage("PDFTest2.xlsx");
        //    PdfPageSettings pageSettings = new PdfPageSettings();
        //    pageSettings.ShowGridLines = true;
        //    pageSettings.PageSize = PdfPageSize.A4;
        //    pageSettings.Orientation = Orientations.Portrait;
        //    pageSettings.Margins = PdfMargins.Normal;
        //    pageSettings.ShowGridLines = true;
        //    //Debug Flags
        //    pageSettings.Debug = true;
        //    pageSettings.PrintAsText = true;

        //    ExcelPdf pedeef = new ExcelPdf(p.Workbook.Worksheets.First(), pageSettings);
        //    pedeef.CreatePdf("c:\\epplustest\\pdf\\EmojiTest.pdf");
        //}

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
            PdfCatalog catlog = new PdfCatalog(pageSettings, ws, "C:\\epplustest\\pdf\\FullPageTest58.pdf");

            //line breaks
            //wrap comments
            //text placement
            //alignment
            //vertical
        }

        [TestMethod]
        public void PerfTest()
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
            PdfCatalog catalog = new PdfCatalog(pageSettings, ws, "C:\\epplustest\\pdf\\OutputTest1.3.pdf");
            Console.WriteLine($"workbook exported time elapsed: {sw.ElapsedMilliseconds} ms");
        }


            [TestMethod]
        // works as expected.
        //[DataRow("PDFTest.xlsx", "C:\\epplustest\\pdf\\FullPageTest56.pdf", "Sheet1")]
        //[DataRow("Aico_0105_S_ALR_87011990_AICO_ASSET_ITE_2025-04_BS.xlsx", "C:\\epplustest\\pdf\\OutputTest1.1.pdf", "SAP Data")]

        // Output file: OutputTest1.2.pdf
        // 1. Minus signs alignment in cells differs from Excel. ------------------------------------------------ Comment: Currently no support for number formats. Requires implementing number formats.
        // 2. Dimension seems to differ from Excel, Excel stops at row 75, EPPlus goes to row 89. --------------- Fixed
        // 3. Row headings are sligthly wider in EPPlus than in Excel. ------------------------------------------ Fixed
        //[DataRow("Aico_0105_S_ALR_87011990_AICO_ASSET_ITE_2025-04_BS.xlsx", "C:\\epplustest\\pdf\\OutputTest1.2.pdf", "Summary")]
        // works as expected
        //[DataRow("Aico WiP 120180 FBL3N for 0110 in 2025-04.xlsx", "C:\\epplustest\\pdf\\OutputTest1.4.pdf", "Technical")]
        //[DataRow("Aico KKS1 Variance Calculation for 0105 in 2025-04 (25_4_2025 15_43_40) .xlsx", "C:\\epplustest\\pdf\\OutputTest1.5.pdf", "Technical")]

        // Output file: OutputTest1.6.pdf
        // 1. Merged cells not working ------------------------------------ Fixed. Comment Merged cells was fine, it was borders being rendered inside merged cells.
        // 2. Pattern fills looks differnt, in some cases not working -----
        // 3. Rotation of text in cells not working (the dates). ----------
        // [DataRow("R05.xlsx", "C:\\epplustest\\pdf\\OutputTest1.6.pdf", "R05 Arbeitseinteilung")]
        //[DataRow("R05 - Copy.xlsx", "C:\\epplustest\\pdf\\OutputTest1.6.pdf", "R05 Arbeitseinteilung")]
        [DataRow("PatternStyles.xlsx", "C:\\epplustest\\pdf\\OutputTest1.8.pdf", "Sheet1")]
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

            PdfCatalog catalog = new PdfCatalog(pageSettings, ws, outputPath);
        }

        /// <summary>
        /// This test we can make somehing of.
        /// </summary>
        //[TestMethod]
        //public void CalculatePages()
        //{
        //    using var p = new ExcelPackage();
        //    var ws = p.Workbook.Worksheets.Add("Sheet 1");
        //    PdfPageSettings pageSettings = new PdfPageSettings();
        //    pageSettings.ShowHeadings = true;
        //    PdfWorksheet pws = new PdfWorksheet();
        //    pws.Worksheet = ws;
        //    pws.ZeroCharWidth = PdfWorksheet.GetThemeFont0Width(ws);
        //    pws.ToRow = 256;
        //    PdfRange range = new PdfRange();
        //    range.TotalWidth = 800;
        //    range.TotalHeight = 1600;

        //    var result = PdfLayout.GetNumberOfPages(pageSettings, pws, ref range);
        //}

        /// <summary>
        /// Old Test
        /// </summary>
        //[TestMethod]
        //public void TestWrapText()
        //{
        //    using var p = OpenTemplatePackage("PDFTest.xlsx");
        //    var cell = p.Workbook.Worksheets[0].Cells["P118"];

        //    List<ITextFragmentBase> TextFragments = GetTextFragments(cell.RichText).Cast<ITextFragmentBase>().ToList();


        //    var layout = OpenTypeFonts.GetTextLayoutEngineForFont((IFontFormatBase)TextFragments[0].RichTextOptions);

        //    var TextLines = layout.WrapRichTextLineCollection(TextFragments, 51d);
        //}

        //private static List<TextFragment> GetTextFragments(ExcelRichTextCollection RichTextCollection, PdfCellStyle cellStyle = null)
        //{
        //    var textFragments = new List<TextFragment>();
        //    bool bold = false, italic = false, underline = false, strike = false;
        //    ExcelUnderLineType underLineType = ExcelUnderLineType.None;
        //    if (cellStyle != null && cellStyle.dxfFont != null)
        //    {
        //        bold = cellStyle.dxfFont.Bold != null ? (bool)cellStyle.dxfFont.Bold : false;
        //        italic = cellStyle.dxfFont.Italic != null ? (bool)cellStyle.dxfFont.Italic : false;
        //        strike = cellStyle.dxfFont.Strike != null ? (bool)cellStyle.dxfFont.Strike : false;
        //        underline = cellStyle.dxfFont.Underline != null;
        //        underLineType = cellStyle.dxfFont.Underline != null ? (ExcelUnderLineType)cellStyle.dxfFont.Underline : ExcelUnderLineType.None;
        //    }
        //    for (int i = 0; i < RichTextCollection.Count; i++)
        //    {
        //        var rt = RichTextCollection[i];
        //        var textFrag = new TextFragment();
        //        textFrag.Text = rt.Text;

        //        textFrag.Font.Family = rt.FontName;
        //        textFrag.Font.Size = rt.Size;

        //        textFrag.RichTextOptions.Bold = rt.Bold || bold;
        //        textFrag.RichTextOptions.Italic = rt.Italic || italic;
        //        //underline
        //        //none   : 12
        //        //single : 13
        //        //Double : 4
        //        //accouting does not exsist
        //        textFrag.RichTextOptions.UnderlineType = 12;
        //        textFrag.RichTextOptions.UnderlineType = rt.UnderLineType == ExcelUnderLineType.Single ? 13 : textFrag.RichTextOptions.UnderlineType;
        //        textFrag.RichTextOptions.UnderlineType = rt.UnderLineType == ExcelUnderLineType.Double ? 4 : textFrag.RichTextOptions.UnderlineType;
        //        textFrag.RichTextOptions.StrikeType = rt.Strike || strike ? 2 : 1;
        //        textFrag.RichTextOptions.SuperScript = rt.VerticalAlign == ExcelVerticalAlignmentFont.Superscript;
        //        textFrag.RichTextOptions.SubScript = rt.VerticalAlign == ExcelVerticalAlignmentFont.Subscript;
        //        textFrag.RichTextOptions.FontColor = rt.Color;

        //        //Should no longer be neccesary
        //        //textFrag.Font.Style = (textFrag.RichTextOptions.Bold ? MeasurementFontStyles.Bold : 0) |
        //        //                      (textFrag.RichTextOptions.Italic ? MeasurementFontStyles.Italic : 0) |
        //        //                      (textFrag.RichTextOptions.UnderlineType != 12 ? MeasurementFontStyles.Underline : 0) |
        //        //                      (textFrag.RichTextOptions.StrikeType > 1 ? MeasurementFontStyles.Strikeout : 0);


        //        textFragments.Add(textFrag);
        //    }

        //    return textFragments;
        //}
    }
}
