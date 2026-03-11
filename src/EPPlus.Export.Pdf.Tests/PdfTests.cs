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
         * Cells: Cutting Text in general. Check cells at text length
         * Cells: Display errors option
         * Cells: Icons
         * Cells: Pictures in cells
         * Cells: Remove stroke from solid fill and adjust size and position of cell more precise and only use fill command.
         * 
         * Merged Cells: Merged cells with no color set will have the color set to white to hide gridlines. They should not have color set to keep tranparency if we want to include background.
         * 
         * Tables: Table Implementation
         * Tables: Fix order of applying table styling. Right now Table is used if it exsists, but cell styling should be prioritised and table should be ignored if cell has styluing.
         * 
         * Patterns: Adjust and make patterns look better. Fixed: DarkUp, DarkDown
         * 
         * Gradients: Make diamond gradients instead of radial gradtient in from corner and center gradient
         * 
         * Text: Bold (Can now do back up bold if bold font not found)
         * Text: Italic (Can now do back up italic if italic font not found)
         * Text: Underline (Basic underline done, need double and single and double accounting)
         * Text: Strikethrough (Done, small adjustments to line to match excel perhaps)
         * Text: Superscript (Done, might need adjustments such as moving it towrds top of cell more)
         * Text: Subscript (Done, might need adjustments. Excel adjusts cell height and moves subscript lower inside the cell)
         * Text: Equations
         * Text: Shrink To Fit
         * Text: Broken text hide when overlapping other cells with text
         * Text: Center vertical text
         * Text: Render text to image from fonts that are not allowed to be embedded?
         * 
         * Layout: Respect Page Breaks
         * Layout: Center on page Horizontal and vertical
         * Layout: Scaling
         * Layout: Fit to number of pages
         * Layout: Row and column headings
         * Layout: Background (Excel does not add this to pdf it seems)
         * Layout: Comments and Notes
         * Layout: Excel Workbook to pdf
         * Layout: Selected worksheets to pdf
         * Layout: Cell range to pdf
         * Layout: Remove empty last page
         * Layout: Calculate width and height of cells more correctly
         * Layout: Fix Mask away stuff outside margins being broken
         * Layout: clipping rekt should exapnd with text longer than its own cell untiln reaching margin area text outside should create a new page. It also should affect gridlines
         * 
         * Borders: Adjust and make border look better
         * 
         * Pivot Table: Pivot table implementation
         * 
         * Drawings: Shapes
         * Drawings Pictures
         * Drawings: Charts
         * Drawings: 3D models
         * 
         * PDF: Compress parts
         */



        [TestMethod]
        public void TestWritePdf()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            //using var p = OpenTemplatePackage("PdfGrids\\PdfTextTest.xlsx");
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
            //Debug Flags
            pageSettings.Debug = true;
            pageSettings.PrintAsText = true;

            ExcelPdf pedeef = new ExcelPdf(ws, pageSettings);
            pedeef.CreatePdf("c:\\epplustest\\pdf\\FullPageTest44.pdf");
        }
    }
}
