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
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using EPPlus.Export.Pdf;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Export.Pdf.PdfSettings.PdfPageSizes;

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
         * Layout: Comments and Notes
         * Layout: Print titles
         * Layout: Background (Excel does not add this to pdf it seems) (Images) : Prio 4
         * Layout: Remove empty last page
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
            //Debug Flags
            pageSettings.Debug = true;
            pageSettings.PrintAsText = true;

            ExcelPdf pedeef = new ExcelPdf(ws, pageSettings);
            pedeef.CreatePdf("c:\\epplustest\\pdf\\FullPageTest48.pdf");
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
    }
}
