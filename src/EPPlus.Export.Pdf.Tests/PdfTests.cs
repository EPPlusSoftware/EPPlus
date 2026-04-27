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
using EPPlus.Export.Pdf;
using EPPlus.Export.Pdf.PdfCatalog;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Export.Pdf.PdfSettings.PdfPageSizes;
using OfficeOpenXml;

namespace EPPlusTest.PDF
{
    [TestClass]
    public class PdfTests : TestBase
    {
        /* BIG PDF TODO
         * This should be turned into tickets on github..
         * 
         * FEATURES:
         * Cells: Display errors option
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

        [TestMethod]
        public void TestWritePdf()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            //using var p = OpenTemplatePackage("PdfGrids\\PdfTextTest.xlsx");
            //using var p = OpenTemplatePackage("PdfGrids\\PdfPageBreakTest.xlsx");
            //using var p = OpenTemplatePackage("PDFTest - Copy (2).xlsx");
            //using var p = OpenTemplatePackage("PdfBorders.xlsx");
            //using var p = OpenTemplatePackage("PdfGrids\\3 2 Page Crazy Cells.xlsx");
            //using var p = OpenTemplatePackage("PdfGrids\\3 2 Page Crazy Cells Merged.xlsx");
            //using var p = OpenTemplatePackage("Gradient.xlsx");
            //using var p = OpenTemplatePackage("PatternFill.xlsx");
            var ws = p.Workbook.Worksheets[0];
            //var ws = p.Workbook.Worksheets[1];
            PdfPageSettings pageSettings = new PdfPageSettings();
            pageSettings.ShowGridLines = true;
            pageSettings.PageSize = PdfPageSize.A4;
            pageSettings.Orientation = Orientations.Portrait;
            pageSettings.Margins = PdfMargins.Normal;
            pageSettings.ShowGridLines = true;
            pageSettings.CenterOnPageHorizontally = true;
            pageSettings.CenterOnPageVertically = true;
            pageSettings.ShowHeadings = true;
            pageSettings.CommentsAndNotes = CommentsAndNotes.AtEndOfSheet;
            //Debug Flags
            pageSettings.Debug = true;
            pageSettings.PrintAsText = true;

            ExcelPdf pedeef = new ExcelPdf(ws, pageSettings);
            pedeef.CreatePdf("c:\\epplustest\\pdf\\FullPageTest49.pdf");
        }

        [TestMethod]
        public void TestWritePdf2()
        {
            using var p = OpenTemplatePackage("PDFTest2.xlsx");
            PdfPageSettings pageSettings = new PdfPageSettings();
            pageSettings.ShowGridLines = true;
            pageSettings.PageSize = PdfPageSize.A4;
            pageSettings.Orientation = Orientations.Portrait;
            pageSettings.Margins = PdfMargins.Normal;
            pageSettings.ShowGridLines = true;
            //Debug Flags
            pageSettings.Debug = true;
            pageSettings.PrintAsText = true;

            ExcelPdf pedeef = new ExcelPdf(p.Workbook.Worksheets.First(), pageSettings);
            pedeef.CreatePdf("c:\\epplustest\\pdf\\EmojiTest.pdf");
        }

        [TestMethod]
        public void ReadPrintAreas()
        {
            //using var p = OpenTemplatePackage("PdfPrintAreas.xlsx");
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            //using var p = OpenTemplatePackage("DoubleBorder.xlsx");
            var ws = p.Workbook.Worksheets[0];
            PdfPageSettings pageSettings = new PdfPageSettings();
            pageSettings.CommentsAndNotes = CommentsAndNotes.AtEndOfSheet;
            pageSettings.Debug = true;
            pageSettings.PrintAsText = true;
            PdfCatalog catlog = new PdfCatalog(pageSettings, ws, "c:\\epplustest\\pdf\\FullPageTest51.pdf");
        }


        [TestMethod]
        public void CalculatePages()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            PdfPageSettings pageSettings = new PdfPageSettings();
            pageSettings.ShowHeadings = true;
            PdfWorksheet pws = new PdfWorksheet();
            pws.Worksheet = ws;
            pws.ZeroCharWidth = PdfWorksheet.GetThemeFont0Width(ws);
            pws.ToRow = 256;
            PdfRange range = new PdfRange();
            range.TotalWidth = 800;
            range.TotalHeight = 1600;

            var result = PdfLayout.GetNumberOfPages(pageSettings, pws, range);
        }
    }
}
