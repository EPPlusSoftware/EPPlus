using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfSettings;
using OfficeOpenXml.PDF.Math;
using System.Collections.Generic;
using System.Linq;
using System;
using OfficeOpenXml.PDF.PdfResources;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfCatalogLayout : PdfTransform
    {
        public PdfCatalogLayout(ExcelRangeBase range, PdfPageSettings pageSettings, Dictionary<string, PdfFontResource> fontResources)
            : base(0, 0, 0, 0)
        {
        }

        public PdfCatalogLayout(ExcelWorkbook workbook, PdfPageSettings pageSettings, Dictionary<string, PdfFontResource> fontResources)
            : base(0, 0, 0, 0)
        {
        }

        public PdfCatalogLayout(ExcelWorksheet worksheet, PdfPageSettings pageSettings, Dictionary<string, PdfFontResource> fontResources, Dictionary<string, PdfPatternResource> patternResources)
            : base(0, 0, 0, 0)
        {
            this.Name = worksheet.Name + " Catalog";
            var WorksheetLayout = AddChild(new PdfWorksheetLayout(worksheet, pageSettings, fontResources, patternResources));
            var PagesLayout = CreatePagesLayoutObject();
            CreatePageLayoutObjects(worksheet, pageSettings, WorksheetLayout, PagesLayout);
            AddCellsToPageLayout(WorksheetLayout, PagesLayout);
            HandleMergedCellsAndDrawings(pageSettings, fontResources, WorksheetLayout, PagesLayout);
            ConvertToPDFCoordiantes(pageSettings, PagesLayout);
            AdjustAndSort(PagesLayout);
            RemoveChild(WorksheetLayout);
            //Save the layout strings to file and then have unity read the files instead of copy and pasting.
            if (pageSettings.Debug)
            {
                string wsLayout = ToHierarchyString();
                string pagesLayout = ToHierarchyString();
                string cellsInPages = ToHierarchyString();
                string mergedCellsInPages = ToHierarchyString();
                string FinalPagesLayout = ToHierarchyString();
            }
        }

        //Create the pages layout.
        private PdfPagesLayout CreatePagesLayoutObject()
        {
            var pages = new PdfPagesLayout(0, 0, 0, 0);
            pages.Name = "Pages";
            AddChild(pages);
            return pages;
        }

        //Create page and content objects.
        private void CreatePageLayoutObjects(ExcelWorksheet worksheet, PdfPageSettings pageSettings, PdfTransform worksheetLayout, PdfPagesLayout pages)
        {
            //Get x cooridiantes to break for new page
            List<double> xBreaks = new List<double>() { 0d };
            double currentWidth = 0, boundsWidth = pageSettings.ContentBounds.Width;
            for (int j = 1; j <= worksheet.Dimension._toCol; j++)
            {
                if (worksheet.Column(j).Hidden) { continue; }
                var width = PdfUnits.ExcelColumnWidthToPoints(worksheet.Column(j).Width);
                if (currentWidth + width >= boundsWidth)
                {
                    xBreaks.Add(currentWidth);
                    boundsWidth = currentWidth + pageSettings.ContentBounds.Width;
                }
                currentWidth += width;
            }
            //Get y cooridiantes to break for new page
            List<double> yBreaks = new List<double>() { 0d };
            double currentHeight = 0, boundsHegiht = pageSettings.ContentBounds.Height;
            for (int i = 1; i <= worksheet.Dimension._toRow; i++)
            {
                if (worksheet.Row(i).Hidden) { continue; }
                var height = PdfUnits.ExcelRowHeightToPoints(worksheet.Row(i).Height);
                if (currentHeight + height >= boundsHegiht)
                {
                    yBreaks.Add(currentHeight);
                    boundsHegiht = currentHeight + pageSettings.ContentBounds.Height;
                }
                currentHeight += height;
            }
            //calculate number of pages needed based on contentBounds and worksheetLayout.Size
            int horizontalPageCount = System.Math.Max(1, (int)System.Math.Ceiling(worksheetLayout.Size.X / pageSettings.ContentBounds.Width));
            int verticalPageCount = System.Math.Max(1, (int)System.Math.Ceiling(worksheetLayout.Size.Y / pageSettings.ContentBounds.Height));
            int totalPages = horizontalPageCount * verticalPageCount;
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
                double x = col * pageSettings.PageSize.WidthPu;
                double y = row * pageSettings.PageSize.HeightPu;
                PdfPageLayout page = new PdfPageLayout(x, y, pageSettings.PageSize.WidthPu, pageSettings.PageSize.HeightPu);
                page.Name = "Page " + (i + 1);
                PdfContentLayout content = new PdfContentLayout(0, 0, pageSettings.ContentBounds);
                content.Name = "Content " + (i + 1);
                page.AddChild(content);
                content.Position = new Vector2(xBreaks[col], yBreaks[row]); //We set position after making content a child of page otherwise positioning breaks and causes errors.
                pages.AddChild(page);
            }
        }

        //Go though all the cells in WorksheetLayout and add them to the overlapping page.
        private void AddCellsToPageLayout(PdfTransform WorksheetLayout, PdfPagesLayout pages)
        {
            var cells = WorksheetLayout.ChildObjects.Where(x => x is PdfCellLayout || x is PdfCellContentLayout || x is PdfCellBorderLayout).ToList();
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
        }

        //Handle merged cells and drawings by checking which pages intersects with them and then make copies for each page.
        private void HandleMergedCellsAndDrawings(PdfPageSettings pageSettings, Dictionary<string, PdfFontResource> fontResources, PdfTransform WorksheetLayout, PdfPagesLayout pages)
        {
            var mcd = WorksheetLayout.ChildObjects.Where(x => x is PdfMergedCellLayout || x is PdfCellContentLayout || x is PdfDrawingLayout).ToList();
            foreach (var mergedCell in mcd)
            {
                var m = mergedCell as PdfMergedCellLayout;
                var c = mergedCell as PdfCellContentLayout;
                var d = mergedCell as PdfDrawingLayout;
                var bounds = mergedCell.GetGlobalBoundingbox();
                foreach (var page in pages.ChildObjects)
                {
                    if (PdfTransform.Intersects(bounds, page.ChildObjects[0].GetGlobalBoundingbox()))
                    {
                        if (m is PdfMergedCellLayout)
                        {
                            var copy = new PdfMergedCellLayout(m.cell, m.LocalPosition.X, m.LocalPosition.Y, m.Size.X, m.Size.Y, m.LocalScale.X, m.LocalScale.Y, m.LocalRotation, WorksheetLayout);
                            copy.Name = m.Name;
                            copy.Z = m.Z;
                            page.ChildObjects[0].AddChild(copy);
                        }
                        else if (c is PdfCellContentLayout)
                        {
                            var copy = new PdfCellContentLayout(c.cell, pageSettings, c.LocalPosition.X, c.LocalPosition.Y, c.Size.X, c.Size.Y, c.LocalScale.X, c.LocalScale.Y, c.LocalRotation, WorksheetLayout, fontResources);
                            copy.Name = c.Name;
                            copy.Z = c.Z;
                            page.ChildObjects[0].AddChild(copy);
                        }
                        else if (d is PdfDrawingLayout) //NOT IMPLEMENTED
                        {
                            var copy = new PdfDrawingLayout(null, d.LocalPosition.X, d.LocalPosition.Y, d.Size.X, d.Size.Y);
                            page.ChildObjects[0].AddChild(copy);
                        }
                    }
                }
            }
        }

        //Restore the positions of the content, move content children to page and remove content object.
        private void ConvertToPDFCoordiantes(PdfPageSettings pageSettings, PdfPagesLayout pages)
        {
            foreach (PdfPageLayout page in pages.ChildObjects)
            {
                page.ChildObjects[0].LocalPosition = new Vector2(pageSettings.Margins.LeftPu, pageSettings.Margins.TopPu);
                var contentObjects = page.ChildObjects[0].ChildObjects.ToList();
                foreach (var child in contentObjects)
                {
                    page.AddChild(child);
                    child.LocalPosition = new Vector2(child.LocalPosition.X, pageSettings.PageSize.HeightPu - System.Math.Abs(child.LocalPosition.Y));
                }
                page.RemoveChild(page.ChildObjects[0]);
                page.GenerateGridLines();
                page.ChildObjects.RemoveAll(x => x.Name.Contains("*")); //Remove all content with * in its name. Better approach would be to not add them at all, But they are needed for grid lines.
            }
        }

        //Make final adjustments and sort children for drawing order.
        private void AdjustAndSort(PdfPagesLayout pages)
        {
            foreach (PdfPageLayout page in pages.ChildObjects)
            {
                //Make adjustments
                foreach (var child in page.ChildObjects)
                {
                    if (child is PdfCellLayout cellLayout)
                    {
                        cellLayout.AdjustForGridLines();
                    }
                    if (child is PdfCellContentLayout contentLayout)
                    {
                        contentLayout.CreatetClippingRect(page.ChildObjects);
                    }
                }
                //Sort by Z ascending and the by Name descending
                page.ChildObjects.Sort((a, b) =>
                {
                    int cmp = a.Z.CompareTo(b.Z);
                    if (cmp == 0)
                        return string.Compare(b.Name, a.Name, StringComparison.OrdinalIgnoreCase);
                    return cmp;
                });
            }
        }
    }
}
