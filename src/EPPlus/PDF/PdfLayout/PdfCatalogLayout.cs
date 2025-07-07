using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfSettings;
using System.Collections.Generic;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfCatalogLayout : PdfTransform
    {
        internal PdfWorksheetLayout wl;
        internal PdfPageSettings settings;
        internal PdfContentBounds bounds;

        public PdfCatalogLayout(PdfWorksheetLayout worksheet, PdfPageSettings pageSettings, PdfContentBounds bounds)
            : base(0, 0, 0, 0)
        {
            wl = worksheet;
            settings = pageSettings;
            this.bounds = bounds;

            //Initialize pageLayouts array by calculating the range for each page.

        }
    }
}



/*
WorksheetLayout
    PageLayout
        HeaderFooterLayout
        ContentLayout //use margins to calculate this
            DrawingsLayout
            CellsLayout
                CellContent // need some sort of cell margins to set posiiton of contents. 
 

1. layout every cell from dimensions in global worksheet layout
2. check page scaling
3. calculate cells to fit on each page
4. Adjust for margins and centering
 */