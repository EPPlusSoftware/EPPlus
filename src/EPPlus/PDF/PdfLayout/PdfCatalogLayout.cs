using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfSettings;
using System.Collections.Generic;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfCatalogLayout : PdfTransform
    {
        internal PdfPageSettings settings;
        internal PdfContentBounds bounds;

        public PdfCatalogLayout(ExcelWorkbook workbook, PdfPageSettings pageSettings, PdfContentBounds bounds)
            : base(0, 0, 0, 0)
        {
        }

        public PdfCatalogLayout(ExcelWorksheet worksheet, PdfPageSettings pageSettings, PdfContentBounds bounds)
            : base(0, 0, 0, 0)
        {
            this.settings = pageSettings;
            this.bounds = bounds;
            var WorksheetLayout = AddChild(new PdfWorksheetLayout(worksheet));
            double x = 0;
            double y = 0;
            var page1Content = AddChild(new PdfContentLayout(x, y, bounds));
            foreach(var child in WorksheetLayout.ChildObjects)
            {
                //need to have a contentLayout
                //check intersect
                //move child to contentLayout
            }
        }

        public PdfCatalogLayout(ExcelRangeBase range, PdfPageSettings pageSettings, PdfContentBounds bounds)
            : base(0, 0, 0, 0)
        {
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