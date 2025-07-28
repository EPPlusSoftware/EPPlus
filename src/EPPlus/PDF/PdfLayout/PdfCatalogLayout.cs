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

            //calculate number of pages needed based on contentBounds and worksheetLayout.Size
            var horizontalPages = WorksheetLayout.Size.X / bounds.Width;
            var verticalPages = WorksheetLayout.Size.Y / bounds.Height;
            int horizontalPageCount = System.Math.Max(1, (int)System.Math.Ceiling(horizontalPages));
            int verticalPageCount = System.Math.Max(1, (int)System.Math.Ceiling(verticalPages));
            int totalPages = horizontalPageCount * verticalPageCount;
            var pages = new PdfPagesLayout(0, 0, 0, 0);
            for (int i = 0; i < totalPages; i++)
            {
                int row, col;
                if (settings.PageOrders == PageOrders.DownThenOver)
                {
                    col = i / verticalPageCount;
                    row = i % verticalPageCount;
                }
                else //(settings.PageOrders == PageOrders.OverThenDown)
                {
                    col = i % horizontalPageCount;
                    row = i / horizontalPageCount;
                }
                double x = col * bounds.Width;
                double y = row * bounds.Height;
                pages.AddChild(new PdfContentLayout(x, y, bounds));
            }
            while (WorksheetLayout.ChildObjects.Count > 0)
            {
                foreach (var cell in WorksheetLayout.ChildObjects)
                {
                    foreach (var page in pages.ChildObjects)
                    {
                        if (PdfTransform.IntersectsFully(page.GetGlobalBoundingbox(), cell.GetGlobalBoundingbox()))
                        {
                            page.AddChild(cell);
                        }
                    }
                    //if cell is not fully covered, move it to the next page and then set new width/height for page. bounds should be the max size not actual page size. we can then set size to be bounds after iterating cells.
                }
            }
            //go into pages and create pageLayout children that contains the contentLayout
        }

        public PdfCatalogLayout(ExcelRangeBase range, PdfPageSettings pageSettings, PdfContentBounds bounds)
            : base(0, 0, 0, 0)
        {
        }
    }
}



/*
WorksheetLayout
PagesLayout
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