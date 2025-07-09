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
            var pages = AddChild(new PdfTransform(0, 0, 0, 0));
            pages.AddChild(new PdfContentLayout(x, y, bounds));

            while (WorksheetLayout.ChildObjects.Count > 0)
            {
                foreach (var cell in WorksheetLayout.ChildObjects)
                {
                    foreach (var page in pages.ChildObjects)
                    {
                        if (PdfTransform.IntersectsFully(page.GetGlobalBoundingbox(), cell.GetGlobalBoundingbox()))
                        {

                        }
                    }
                    //if cell is not fully covered, move it to the next page and then set new width/height for page. bounds should be the max size not actual page size. we can then set size to be bounds after iterating cells.


                    if (settings.PageOrders == PageOrders.DownThenOver)
                    {
                        //add page in y coord first
                    }
                    else if (settings.PageOrders == PageOrders.OverThenDown)
                    {
                        //add page in x coord first
                    }
                }
            }





            //foreach(var child in WorksheetLayout.ChildObjects)
            //{
            //    bool childInPage = false;
            //    foreach(var page in pages.ChildObjects)
            //    {
            //        if(child.Intersects(child.GetGlobalBoundingbox(), page.GetGlobalBoundingbox()))
            //        {
            //            child.Parent = page;
            //            childInPage = true;
            //        }
            //    }
            //    if(childInPage = false)
            //    {

            //    }
            //    //need to have a contentLayout
            //    //check intersect
            //    //move child to contentLayout
            //    //here we can add pages in over then down order or down the over order.
            //}
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