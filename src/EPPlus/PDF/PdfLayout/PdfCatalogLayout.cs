using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfSettings;
using OfficeOpenXml.PDF.Math;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

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
            var WorksheetLayout = AddChild(new PdfWorksheetLayout(worksheet, pageSettings));
            //calculate number of pages needed based on contentBounds and worksheetLayout.Size
            var horizontalPages = WorksheetLayout.Size.X / bounds.Width;
            var verticalPages = WorksheetLayout.Size.Y / bounds.Height;
            int horizontalPageCount = System.Math.Max(1, (int)System.Math.Ceiling(horizontalPages));
            int verticalPageCount = System.Math.Max(1, (int)System.Math.Ceiling(verticalPages));
            int totalPages = horizontalPageCount * verticalPageCount;
            var pages = new PdfPagesLayout(0, 0, 0, 0);
            pages.Name = "Pages";
            AddChild(pages);
            string wsLayout = ToHierarchyString();
            //Create lists for content objects starting positions.
            List<double> xBreaks = new List<double>() { 0d };
            List<double> yBreaks = new List<double>() { 0d };
            double h = 0;
            double w = 0;
            int pw = 1;
            int ph = 1;
            double bh = bounds.Height;
            double bw = bounds.Width;
            for (int i = 1; i <= worksheet.Dimension._toRow; i++)
            {
                if (worksheet.Row(i).Hidden) { continue; }
                var height = worksheet.Row(i).Height;
                var cell = worksheet.Cells[i, 1];
                if (h+height >= bh)
                {
                    ph++;
                    yBreaks.Add(h);
                    bh = h + bounds.Height;
                }
                h += height;
            }
            for (int j = 1; j <= worksheet.Dimension._toCol; j++)
            {
                if (worksheet.Column(j).Hidden) { continue; }
                var width = PdfUnits.ExcelColumnWidthToPoints(worksheet.Column(j).Width);
                if (w+width >= bw)
                {
                    pw++;
                    xBreaks.Add(w);
                    bw = w + bounds.Width;
                }
                w += width;
            }
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
                content.Position = new Vector2(xBreaks[col], yBreaks[row]);
                pages.AddChild(page);
            }
            string pagesLayout = ToHierarchyString();
            //Go though all the cells in WorksheetLayout and add them to the overlapping page.
            var cells = WorksheetLayout.ChildObjects.Where(x=>x is PdfCellLayout).ToList();
            foreach (var cell in cells)
            {
                var cellBounds = cell.GetGlobalBoundingbox();
                foreach (var page in pages.ChildObjects)
                {
                    var contentbounds = page.ChildObjects[0].GetGlobalBoundingbox();
                    if (PdfTransform.IntersectsFully(contentbounds, cellBounds))
                    {
                        page.ChildObjects[0].AddChild(cell);
                        break;
                    }
                }
            }
            string cellsInPages = ToHierarchyString();
            //handle merged cells
            var mergedCells = WorksheetLayout.ChildObjects.Where(x => x is PdfMergedCellLayout).ToList();
            foreach (var mc in mergedCells)
            {
                var mcBounds = mc.GetGlobalBoundingbox();
                foreach(var page in pages.ChildObjects)
                {
                    if(PdfTransform.Intersects(mcBounds, page.ChildObjects[0].GetGlobalBoundingbox()))
                    {
                        var copy = new PdfMergedCellLayout(null,pageSettings, mc.LocalPosition.X, mc.LocalPosition.Y, mc.Size.X, mc.Size.Y, mc.LocalScale.X, mc.LocalScale.Y, mc.LocalRotation, WorksheetLayout);
                        copy.Name = mc.Name;
                        page.ChildObjects[0].AddChild(copy);
                    }
                }
            }
            string mergedCellsInPages = ToHierarchyString();
            //handle drawings
            var drawings = WorksheetLayout.ChildObjects.Where(x => x is PdfDrawingLayout).ToList();
            foreach (var d in drawings)
            {
                var dBounds = d.GetGlobalBoundingbox();
                foreach (var page in pages.ChildObjects)
                {
                    if (PdfTransform.Intersects(dBounds, page.ChildObjects[0].GetGlobalBoundingbox()))
                    {
                        var copy = new PdfDrawingLayout(null, d.LocalPosition.X, d.LocalPosition.Y, d.Size.X, d.Size.Y);
                        page.ChildObjects[0].AddChild(copy);
                    }
                }
            }
            string drawingsCellsInPages = ToHierarchyString();
            //Restore the positions of the pages
            foreach (var page in pages.ChildObjects)
            {
                page.ChildObjects[0].LocalPosition = new Vector2(settings.Margins.LeftPu, settings.Margins.TopPu);
            }
            RemoveChild(WorksheetLayout);
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