using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfSettings;
using OfficeOpenXml.PDF.Math;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.PDF.PdfFontData;
using System.Runtime.InteropServices;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfCatalogLayout : PdfTransform
    {
        public PdfCatalogLayout(ExcelRangeBase range, PdfPageSettings pageSettings, PdfContentBounds bounds, Dictionary<string, PdfFontResource> fontResources)
            : base(0, 0, 0, 0)
        {
        }

        public PdfCatalogLayout(ExcelWorkbook workbook, PdfPageSettings pageSettings, PdfContentBounds bounds, Dictionary<string, PdfFontResource> fontResources)
            : base(0, 0, 0, 0)
        {
        }

        public PdfCatalogLayout(ExcelWorksheet worksheet, PdfPageSettings pageSettings, PdfContentBounds bounds, Dictionary<string, PdfFontResource> fontResources)
            : base(0, 0, 0, 0)
        {
            //Position = new Vector2(0d, pageSettings.PageSize.HeightPu);
            this.Name = worksheet.Name + " Catalog";
            var WorksheetLayout = AddChild(new PdfWorksheetLayout(worksheet, pageSettings, bounds, fontResources));
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
                var height = PdfUnits.ExcelRowHeightToPoints(worksheet.Row(i).Height);
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
                if (pageSettings.PageOrders == PageOrders.DownThenOver)
                {
                    col = i / verticalPageCount;
                    row = i % verticalPageCount;
                }
                else //(settings.PageOrders == PageOrders.OverThenDown)
                {
                    col = i % horizontalPageCount;
                    row = i / horizontalPageCount;
                }
                double px = col * pageSettings.PageSize.WidthPu;
                double py = row * pageSettings.PageSize.HeightPu;
                PdfPageLayout page = new PdfPageLayout(px, py, pageSettings.PageSize.WidthPu, pageSettings.PageSize.HeightPu);
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
            var cells = WorksheetLayout.ChildObjects.Where(x=>x is PdfCellLayout || x is PdfCellContentLayout || x is PdfCellBorderLayout).ToList();
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
            var mergedCells = WorksheetLayout.ChildObjects.Where(x => x is PdfMergedCellLayout || x is PdfCellContentLayout).ToList();
            foreach (var mergedCell in mergedCells)
            {
                var mcl = mergedCell as PdfMergedCellLayout;
                var mcc = mergedCell as PdfCellContentLayout;
                var mcBounds = mergedCell.GetGlobalBoundingbox();
                foreach(var page in pages.ChildObjects)
                {
                    if(PdfTransform.Intersects(mcBounds, page.ChildObjects[0].GetGlobalBoundingbox()))
                    {
                        if (mcl is PdfMergedCellLayout)
                        {
                            var copy = new PdfMergedCellLayout(mcl.cell, mcl.LocalPosition.X, mcl.LocalPosition.Y, mcl.Size.X, mcl.Size.Y, mcl.LocalScale.X, mcl.LocalScale.Y, mcl.LocalRotation, WorksheetLayout);
                            copy.Name = mcl.Name;
                            copy.Z = mcl.Z;
                            page.ChildObjects[0].AddChild(copy);
                        }
                        else if(mcc is  PdfCellContentLayout)
                        {
                            var copy = new PdfCellContentLayout(mcc.cell, pageSettings, mcc.LocalPosition.X, mcc.LocalPosition.Y, mcc.Size.X, mcc.Size.Y, mcc.LocalScale.X, mcc.LocalScale.Y, mcc.LocalRotation, WorksheetLayout, fontResources);
                            copy.Name = mcc.Name;
                            copy.Z = mcc.Z;
                            page.ChildObjects[0].AddChild(copy);
                        }
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
            //Restore the positions of the content, move content children to page and remove content object.
            foreach (PdfPageLayout page in pages.ChildObjects)
            {
                page.ChildObjects[0].LocalPosition = new Vector2(pageSettings.Margins.LeftPu, pageSettings.Margins.TopPu);
                var contentObjects = page.ChildObjects[0].ChildObjects.ToList();
                foreach (var child in contentObjects)
                {
                    page.AddChild(child);
                    child.LocalPosition = new Vector2(child.LocalPosition.X, pageSettings.PageSize.HeightPu - System.Math.Abs(child.LocalPosition.Y));
                    if (child is PdfCellContentLayout contentLayout)
                    {
                        contentLayout.AdjustClipping(pageSettings.PageSize.HeightPu);
                    }
                }
                page.RemoveChild(page.ChildObjects[0]);
                //page.Range = worksheet.Cells[ page.ChildObjects[0].Name + ":" + page.ChildObjects[page.ChildObjects.Count - 1].Name];
                page.GenerateGridLines(pageSettings, worksheet);
                page.ChildObjects.RemoveAll(x => x.Name.Contains("*"));
            }
            foreach (PdfPageLayout page in pages.ChildObjects)
            {
                foreach (var child in page.ChildObjects)
                {
                    if( child is PdfCellLayout cellLayout)
                    {
                        cellLayout.Adjust();
                    }
                }
                page.ChildObjects.Sort((a, b) => a.Z.CompareTo(b.Z));
            }
            RemoveChild(WorksheetLayout);
            string FinalPagesLayout = ToHierarchyString();
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