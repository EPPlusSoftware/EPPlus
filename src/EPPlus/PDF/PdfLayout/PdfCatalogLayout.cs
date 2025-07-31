using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfSettings;
using OfficeOpenXml.PDF.Math;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfCatalogLayout : PdfTransform
    {
        internal PdfPageSettings settings;
        internal PdfContentBounds bounds;

        public PdfCatalogLayout(ExcelRangeBase range, PdfPageSettings pageSettings, PdfContentBounds bounds)
            : base(0, 0, 0, 0)
        {
        }

        public PdfCatalogLayout(ExcelWorkbook workbook, PdfPageSettings pageSettings, PdfContentBounds bounds)
            : base(0, 0, 0, 0)
        {
        }

        public PdfCatalogLayout(ExcelWorksheet worksheet, PdfPageSettings pageSettings, PdfContentBounds bounds)
            : base(0, 0, 0, 0)
        {
            this.Name = worksheet.Name + " Catalog";
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
            pages.Name = "Pages";
            AddChild(pages);
            //Create the new pages and place them in a grid.
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
                double px = col * settings.PageSize.WidthPu;
                double py = row * settings.PageSize.HeightPu;
                PdfPageLayout page = new PdfPageLayout(px, py, settings.PageSize.WidthPu, settings.PageSize.HeightPu);
                page.Name = "Page " + (i + 1);
                double cx = col * bounds.Width;
                double cy = row * bounds.Height;
                //var contentLocalX = cx - px;
                //var contentLocalY = cy - py;
                PdfContentLayout content = new PdfContentLayout(0, 0, bounds);
                content.Name = "Content " + (i + 1);
                page.AddChild(content);
                content.Position = new Vector2(cx, cy);
                pages.AddChild(page);
            }
            //Go though all the cells in WorksheetLayout and add them to the overlapping page.
            var cells = WorksheetLayout.ChildObjects.ToList();
            foreach (var cell in cells)
            {
                foreach (var page in pages.ChildObjects)
                {
                    bool move = false;
                    var cellBounds = cell.GetGlobalBoundingbox();
                    foreach (var p in pages.ChildObjects)
                    {
                        var contentBounds = p.ChildObjects[0].GetGlobalBoundingbox();
                        //If the cell is completly inside a page content. Make that cell a child of the page.
                        if (PdfTransform.IntersectsFully(contentBounds, cellBounds))
                        {
                            move = true;
                            p.ChildObjects[0].AddChild(cell);
                            break;
                        }
                    }
                    if (!move)
                    {
                        move = false;
                        //If the cell is only partially inside a page content. Move the page to overlap the cell. This will make pages overlap, but we will fix this later
                        var neighborContent = GetRightBottomAndDiagonalPages(page, pages.ChildObjects, horizontalPageCount, verticalPageCount, settings.PageOrders);
                        foreach (var neighbor in neighborContent)
                        {
                            var neighborBounds = neighbor.ChildObjects[0].GetGlobalBoundingbox();
                            if (PdfTransform.Intersects(cellBounds, neighborBounds))
                            {
                                // Temporarily move the neighbor page to align with the cell
                                var dx = cellBounds.X - neighborBounds.X;
                                var dy = cellBounds.Y - neighborBounds.Y;
                                neighbor.ChildObjects[0].Translate(dx, dy);
                                neighbor.ChildObjects[0].AddChild(cell);
                                break;
                            }
                        }
                    }
                }
            }
            //Restore the positions of the pages
            foreach (var page in pages.ChildObjects)
            {
                page.ChildObjects[0].LocalPosition = new Vector2(settings.Margins.LeftPu, settings.Margins.TopPu);
            }
        }

        List<PdfPageLayout> GetRightBottomAndDiagonalPages(PdfTransform currentPage, List<PdfTransform> allPages, int hPages, int vPages, PageOrders pageOrder)
        {
            int index = allPages.IndexOf(currentPage);
            if (index == -1)
                return new List<PdfPageLayout>();

            int row, col;

            if (pageOrder == PageOrders.DownThenOver)
            {
                col = index / vPages;
                row = index % vPages;
            }
            else //(settings.PageOrders == PageOrders.OverThenDown)
            {
                col = index % hPages;
                row = index / hPages;
            }
            var neighbors = new List<PdfPageLayout>();
            // Right
            if (col + 1 < hPages)
            {
                int i = (pageOrder == PageOrders.DownThenOver) ? (col + 1) * vPages + row : row * hPages + (col + 1);
                neighbors.Add(allPages[i] as PdfPageLayout);
            }
            // Bottom
            if (row + 1 < vPages)
            {
                int i = (pageOrder == PageOrders.DownThenOver) ? col * vPages + (row + 1) : (row + 1) * hPages + col;
                neighbors.Add(allPages[i] as PdfPageLayout);
            }
            // Bottom-right
            if (col + 1 < hPages && row + 1 < vPages)
            {
                int i = (pageOrder == PageOrders.DownThenOver) ? (col + 1) * vPages + (row + 1) : (row + 1) * hPages + (col + 1);
                neighbors.Add(allPages[i] as PdfPageLayout);
            }
            return neighbors;
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