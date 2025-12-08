/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using System.Collections.Generic;
using System.Linq;
using System;
using OfficeOpenXml;
using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Graphics;
using EPPlus.Graphics.Math;
using EPPlus.Graphics.Units;
using OfficeOpenXml.Style.HeaderFooterTextFormat;

namespace EPPlus.Export.Pdf.PdfLayout
{
    internal class PdfCatalogLayout : Transform
    {
        public PdfCatalogLayout(ExcelRangeBase range, PdfPageSettings pageSettings, PdfDictionaries dictionaries)
            : base(0, 0, 0, 0)
        {
        }

        public PdfCatalogLayout(ExcelWorkbook workbook, PdfPageSettings pageSettings, PdfDictionaries dictionaries)
            : base(0, 0, 0, 0)
        {
        }

        public PdfCatalogLayout(ExcelWorksheet worksheet, PdfPageSettings pageSettings, PdfDictionaries dictionaries)
            : base(0, 0, 0, 0)
        {
            Name = worksheet.Name + " Catalog";
            var WorksheetLayout = AddChild(new PdfWorksheetLayout(worksheet, pageSettings, dictionaries));
            string wsLayout = ToHierarchyString();
            var PagesLayout = CreatePagesLayoutObject();
            CreatePageLayoutObjects(worksheet, pageSettings, WorksheetLayout as PdfWorksheetLayout, PagesLayout);
            string pagesLayout = ToHierarchyString();
            AddCellsToPageLayout(WorksheetLayout, PagesLayout);
            string cellsInPages = ToHierarchyString();
            HandleMergedCellsAndDrawings(pageSettings, dictionaries, WorksheetLayout, PagesLayout);
            string mergedCellsInPages = ToHierarchyString();
            ConvertToPDFCoordiantes(pageSettings, PagesLayout);
            string ConvertedCoordinates = ToHierarchyString();
            AdjustAndSort(PagesLayout);
            RemoveChild(WorksheetLayout);
            AddHeaderFooter(worksheet, pageSettings, dictionaries, PagesLayout);
            string FinalPagesLayout = ToHierarchyString();
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
        private void CreatePageLayoutObjects(ExcelWorksheet worksheet, PdfPageSettings pageSettings, PdfWorksheetLayout worksheetLayout, PdfPagesLayout pages)
        {
            //Get x cooridiantes to break for new page
            List<double> xBreaks = new List<double>() { 0d };
            double currentWidth = 0, boundsWidth = pageSettings.ContentBounds.Width;
            for (int j = 1; j <= worksheet.Dimension._toCol; j++)
            {
                if (worksheet.Column(j).Hidden) { continue; }
                var width = UnitConversion.ExcelColumnWidthToPoints(worksheet.Column(j).Width, worksheetLayout.ZeroCharWidth);
                if (currentWidth + width >= boundsWidth)
                {
                    xBreaks.Add(currentWidth);
                    boundsWidth = currentWidth + pageSettings.ContentBounds.Width;
                }
                currentWidth += width;
            }
            //Get y cooridiantes to break for new page
            List<double> yBreaks = new List<double>() { 0d };
            double currentHeight = 0, boundsHegiht = -pageSettings.ContentBounds.Height;
            for (int i = 1; i <= worksheet.Dimension._toRow; i++)
            {
                if (worksheet.Row(i).Hidden) { continue; }
                var height = UnitConversion.ExcelRowHeightToPoints(worksheet.Row(i).Height);
                if (currentHeight - height <= boundsHegiht)
                {
                    yBreaks.Add(currentHeight);
                    boundsHegiht = currentHeight - pageSettings.ContentBounds.Height;
                }
                currentHeight -= height;
            }
            //calculate number of pages needed based on contentBounds and worksheetLayout.Size
            int horizontalPageCount = System.Math.Max(1, (int)System.Math.Ceiling(worksheetLayout.Size.X / pageSettings.ContentBounds.Width));
            int verticalPageCount = System.Math.Max(1, (int)System.Math.Ceiling(System.Math.Abs( worksheetLayout.Size.Y) / pageSettings.ContentBounds.Height));
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
                double y = row * -pageSettings.PageSize.HeightPu;
                PdfPageLayout page = new PdfPageLayout(x, y-pageSettings.PageSize.HeightPu, pageSettings.PageSize.WidthPu, pageSettings.PageSize.HeightPu);
                page.Name = "Page " + (i + 1);
                PdfContentLayout content = new PdfContentLayout(0, 0, pageSettings.ContentBounds);
                content.Name = "Content " + (i + 1);
                page.AddChild(content);
                content.Position = new Vector2(xBreaks[col], yBreaks[row] - pageSettings.ContentBounds.Height); //We set position after making content a child of page otherwise positioning breaks and causes errors.
                pages.AddChild(page);
            }
        }

        //Go though all the cells in WorksheetLayout and add them to the overlapping page.
        private void AddCellsToPageLayout(Transform WorksheetLayout, PdfPagesLayout pages)
        {
            var cells = WorksheetLayout.ChildObjects.Where(x => x is PdfCellLayout || x is PdfCellContentLayout || x is PdfCellBorderLayout).ToList();
            foreach (var cell in cells)
            {
                var cellBounds = cell.GetGlobalBoundingbox();
                foreach (var page in pages.ChildObjects)
                {
                    var contentbounds = page.ChildObjects[0].GetGlobalBoundingbox();
                    if (IntersectsFully(contentbounds, cellBounds))
                    {
                        page.ChildObjects[0].AddChild(cell);
                        break;
                    }
                }
            }
        }

        //Handle merged cells and drawings by checking which pages intersects with them and then make copies for each page.
        private void HandleMergedCellsAndDrawings(PdfPageSettings pageSettings, PdfDictionaries dictionaries, Transform WorksheetLayout, PdfPagesLayout pages)
        {
            var mcd = WorksheetLayout.ChildObjects.Where(x => x is PdfMergedCellLayout || x is PdfCellContentLayout || x is PdfCellBorderLayout || x is PdfDrawingLayout).ToList();
            foreach (var mergedCell in mcd)
            {
                var m = mergedCell as PdfMergedCellLayout;
                var c = mergedCell as PdfCellContentLayout;
                var b = mergedCell as PdfCellBorderLayout;
                var d = mergedCell as PdfDrawingLayout;
                var bounds = mergedCell.GetGlobalBoundingbox();
                foreach (var page in pages.ChildObjects)
                {
                    if (Intersects(bounds, page.ChildObjects[0].GetGlobalBoundingbox()))
                    {
                        if (m is PdfMergedCellLayout)
                        {
                            var copy = new PdfMergedCellLayout(dictionaries, m.cell, m.LocalPosition.X, m.LocalPosition.Y + m.Size.Y, m.Size.X, m.Size.Y, m.LocalScale.X, m.LocalScale.Y, m.LocalRotation, WorksheetLayout);
                            copy.Name = m.Name;
                            copy.Z = m.Z;
                            page.ChildObjects[0].AddChild(copy);
                        }
                        else if (c is PdfCellContentLayout)
                        {
                            var copy = new PdfCellContentLayout(c.cell, pageSettings, c.LocalPosition.X, c.LocalPosition.Y, c.Size.X, c.Size.Y, c.LocalScale.X, c.LocalScale.Y, c.LocalRotation, WorksheetLayout, dictionaries);
                            copy.Name = c.Name;
                            copy.Z = c.Z;
                            page.ChildObjects[0].AddChild(copy);
                        }
                        else if (b is PdfCellBorderLayout)
                        {
                            var copy = new PdfCellBorderLayout(b.cell, b.LocalPosition.X, b.LocalPosition.Y + b.Size.Y, b.Size.X, b.Size.Y, b.LocalScale.X, b.LocalScale.Y, b.LocalRotation, WorksheetLayout);
                            copy.Name = b.Name;
                            copy.Z = b.Z;
                            copy.BorderData = b.BorderData;
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
                    //if (child is ILayout il)
                    //    il.ConvertCoordinates(pageSettings);
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
                        contentLayout.CreateClippingRect(page.ChildObjects);
                    }
                    if (child is PdfMergedCellLayout mergedLayout)
                    {
                        mergedLayout.AdjustForGridLines();
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

        private void AddHeaderFooter(ExcelWorksheet ws, PdfPageSettings settings, PdfDictionaries dictionaries, PdfPagesLayout pages)
        {
            int pageNumber = 1;
            //loop pages and check which header it should use
            for (int i = 0; i < pages.ChildObjects.Count; i++)
            {
                ExcelHeaderFooterTextCollection leftH = null;
                ExcelHeaderFooterTextCollection centerH = null;
                ExcelHeaderFooterTextCollection rightH = null;
                ExcelHeaderFooterTextCollection leftF = null;
                ExcelHeaderFooterTextCollection centerF = null;
                ExcelHeaderFooterTextCollection rightF = null;

                if (ws.HeaderFooter.differentFirst && pageNumber == 1)
                {
                    leftH = ws.HeaderFooter.FirstHeader.LeftAligned;
                    centerH = ws.HeaderFooter.FirstHeader.Centered;
                    rightH = ws.HeaderFooter.FirstHeader.RightAligned;
                    leftF = ws.HeaderFooter.FirstFooter.LeftAligned;
                    centerF = ws.HeaderFooter.FirstFooter.Centered;
                    rightF = ws.HeaderFooter.FirstFooter.RightAligned;
                    if (leftH.Text == "&L" && centerH.Text == "&C" && rightH.Text == "&R" && leftF.Text == "&L" && centerF.Text == "&C" && rightF.Text == "&R")
                    {
                        leftH = ws.HeaderFooter.OddHeader.LeftAligned;
                        centerH = ws.HeaderFooter.OddHeader.Centered;
                        rightH = ws.HeaderFooter.OddHeader.RightAligned;
                        leftF = ws.HeaderFooter.OddFooter.LeftAligned;
                        centerF = ws.HeaderFooter.OddFooter.Centered;
                        rightF = ws.HeaderFooter.OddFooter.RightAligned;
                    }
                }
                else if (ws.HeaderFooter.differentOddEven && pageNumber % 2 == 0)
                {
                    leftH = ws.HeaderFooter.EvenHeader.LeftAligned;
                    centerH = ws.HeaderFooter.EvenHeader.Centered;
                    rightH = ws.HeaderFooter.EvenHeader.RightAligned;
                    leftF = ws.HeaderFooter.EvenFooter.LeftAligned;
                    centerF = ws.HeaderFooter.EvenFooter.Centered;
                    rightF = ws.HeaderFooter.EvenFooter.RightAligned;
                    if (leftH.Text == "&L" && centerH.Text == "&C" && rightH.Text == "&R" && leftF.Text == "&L" && centerF.Text == "&C" && rightF.Text == "&R")
                    {
                        leftH = ws.HeaderFooter.OddHeader.LeftAligned;
                        centerH = ws.HeaderFooter.OddHeader.Centered;
                        rightH = ws.HeaderFooter.OddHeader.RightAligned;
                        leftF = ws.HeaderFooter.OddFooter.LeftAligned;
                        centerF = ws.HeaderFooter.OddFooter.Centered;
                        rightF = ws.HeaderFooter.OddFooter.RightAligned;
                    }
                }
                else
                {
                    leftH = ws.HeaderFooter.OddHeader.LeftAligned;
                    centerH = ws.HeaderFooter.OddHeader.Centered;
                    rightH = ws.HeaderFooter.OddHeader.RightAligned;
                    leftF = ws.HeaderFooter.OddFooter.LeftAligned;
                    centerF = ws.HeaderFooter.OddFooter.Centered;
                    rightF = ws.HeaderFooter.OddFooter.RightAligned;
                }
                var lh = (PdfHeaderFooterLayout)pages.ChildObjects[i].AddChild(new PdfHeaderFooterLayout(leftH, ws, settings, dictionaries, pageNumber, pages.ChildObjects.Count));
                lh.LocalPosition = new Vector2(settings.Margins.LeftPu, settings.PageSize.HeightPu - settings.Margins.HeaderPu);
                lh.AdjustPositionByTextLength('l', 'h');
                var ch = (PdfHeaderFooterLayout)pages.ChildObjects[i].AddChild(new PdfHeaderFooterLayout(centerH, ws, settings, dictionaries, pageNumber, pages.ChildObjects.Count));
                ch.LocalPosition = new Vector2(settings.PageSize.WidthPu / 2d, settings.PageSize.HeightPu - settings.Margins.HeaderPu);
                ch.AdjustPositionByTextLength('c', 'h');
                var rh = (PdfHeaderFooterLayout)pages.ChildObjects[i].AddChild(new PdfHeaderFooterLayout(rightH, ws, settings, dictionaries, pageNumber, pages.ChildObjects.Count));
                rh.LocalPosition = new Vector2(settings.PageSize.WidthPu - settings.Margins.RightPu, settings.PageSize.HeightPu - settings.Margins.HeaderPu);
                rh.AdjustPositionByTextLength('r', 'h');
                var lf = pages.ChildObjects[i].AddChild(new PdfHeaderFooterLayout(leftF, ws, settings, dictionaries, pageNumber, pages.ChildObjects.Count));
                lf.LocalPosition = new Vector2(settings.Margins.LeftPu, settings.Margins.FooterPu);
                var cf = (PdfHeaderFooterLayout)pages.ChildObjects[i].AddChild(new PdfHeaderFooterLayout(centerF, ws, settings, dictionaries, pageNumber, pages.ChildObjects.Count));
                cf.LocalPosition = new Vector2(settings.PageSize.WidthPu / 2d, settings.Margins.FooterPu);
                cf.AdjustPositionByTextLength('c', 'f');
                var rf = (PdfHeaderFooterLayout)pages.ChildObjects[i].AddChild(new PdfHeaderFooterLayout(rightF, ws, settings, dictionaries, pageNumber, pages.ChildObjects.Count));
                rf.LocalPosition = new Vector2(settings.PageSize.WidthPu - settings.Margins.RightPu, settings.Margins.FooterPu);
                rf.AdjustPositionByTextLength('r', 'f');
                pageNumber++;
            }
        }
    }
}
