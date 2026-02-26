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
using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Graphics;
using EPPlus.Graphics.Math;
using EPPlus.Graphics.Units;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.HeaderFooterTextFormat;
using System;
using System.Collections.Generic;
using System.Linq;

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
            CreateFontSubsets(pageSettings, dictionaries.Fonts);
            var PagesLayout = CreatePagesLayoutObject();
            CreatePageLayoutObjects(worksheet, pageSettings, WorksheetLayout as PdfWorksheetLayout, PagesLayout);
            AddCellsToPageLayout(WorksheetLayout, PagesLayout);
            HandleMergedCellsAndDrawings(pageSettings, dictionaries, WorksheetLayout, PagesLayout);
            MoveCellToPageFromContent(pageSettings, PagesLayout);
            ProocessPageAndCells(pageSettings, dictionaries, PagesLayout);
            //ConvertToPDFCoordiantes(pageSettings, PagesLayout, worksheet);
            //AdjustAndSort(PagesLayout, dictionaries);
            RemoveChild(WorksheetLayout);
            AddHeaderFooter(worksheet, pageSettings, dictionaries, PagesLayout);
        }

        private void CreateFontSubsets(PdfPageSettings pageSettings, Dictionary<string, PdfFontResource> fonts)
        {
            foreach (var font in fonts)
            {
                font.Value.CreateSubset();
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
        private void CreatePageLayoutObjects(ExcelWorksheet worksheet, PdfPageSettings pageSettings, PdfWorksheetLayout worksheetLayout, PdfPagesLayout pages)
        {
            //Get x cooridiantes to break for new page
            List<double> xBreaks = new List<double>() { 0d };
            double currentWidth = 0, boundsWidth = pageSettings.ContentBounds.Width;
            for (int j = 1; j <= worksheet.Dimension._toCol; j++)
            {
                if (worksheet.Column(j).Hidden) { continue; }
                var width = UnitConversion.ExcelColumnWidthToPoints(worksheet.Column(j).Width, PdfWorksheetLayout.ZeroCharWidth);
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
                foreach (PdfPageLayout page in pages.ChildObjects)
                {
                    var contentbounds = page.ChildObjects[0].GetGlobalBoundingbox();
                    if (IntersectsFully(contentbounds, cellBounds))
                    {
                        page.AddCell(cell);
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
                            var copy = new PdfMergedCellLayout(dictionaries, m.cell, m.CellStyle, m.LocalPosition.X, m.LocalPosition.Y + m.Size.Y, m.Size.X, m.Size.Y, m.LocalScale.X, m.LocalScale.Y, m.LocalRotation, WorksheetLayout);
                            copy.Name = m.Name;
                            copy.Z = m.Z;
                            page.ChildObjects[0].AddChild(copy);
                        }
                        else if (c is PdfCellContentLayout)
                        {
                            var copy = new PdfCellContentLayout(c.cell, c.CellStyle, pageSettings, c.LocalPosition.X, c.LocalPosition.Y, c.Size.X, c.Size.Y, c.LocalScale.X, c.LocalScale.Y, c.LocalRotation, WorksheetLayout, dictionaries);
                            copy.Name = c.Name;
                            copy.Z = c.Z;
                            page.ChildObjects[0].AddChild(copy);
                        }
                        else if (b is PdfCellBorderLayout)
                        {
                            var copy = new PdfCellBorderLayout(b.cell, null, b.LocalPosition.X, b.LocalPosition.Y + b.Size.Y, b.Size.X, b.Size.Y, b.LocalScale.X, b.LocalScale.Y, b.LocalRotation, WorksheetLayout);
                            copy.Name = b.Name;
                            copy.Z = b.Z;
                            copy.BorderData = b.BorderData;
                            copy.range = b.range;
                            copy.IsMerged = b.IsMerged;
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

        private void MoveCellToPageFromContent(PdfPageSettings pageSettings, PdfPagesLayout pages)
        {
            foreach (PdfPageLayout page in pages.ChildObjects)
            {
                var rowCount = page.ToRow - page.FromRow + 1;
                var colCount = page.ToCol - page.FromCol + 1;
                int col = 0;
                int row = 0;
                page.CreateMap();
                page.ChildObjects[0].LocalPosition = new Vector2(pageSettings.Margins.LeftPu, pageSettings.Margins.TopPu);
                var contentObjects = page.ChildObjects[0].ChildObjects.ToList();
                for (int i = 0; i < contentObjects.Count; i++)
                {
                    var child = contentObjects[i];
                    page.AddChild(child);
                    if (child is IShadingLayout iSl)
                        iSl.UpdateShadingPositionMatrix(pageSettings);
                    if (child is PdfMergedCellLayout m)
                    {
                        var mRowCount = m.address._toRow - m.address._fromRow + 1;
                        var mColCount = m.address._toCol - m.address._fromCol + 1;
                        var innerCol = col;
                        var innerRow = row;
                        for (int r = 0; r < mRowCount; r++)
                        {
                            for (int c = 0; c < mColCount; c++)
                            {
                                page.Map[innerRow, innerCol].Name = ExcelRange.GetColumnLetter(innerCol+1) + (innerRow+1);
                                page.Map[innerRow, innerCol].cell = m;
                                page.Map[innerRow, innerCol].Type = PageMap.CellType.Merged;
                                page.Map[innerRow, innerCol].content = contentObjects.OfType<PdfCellContentLayout>().FirstOrDefault(t => t.Name == m.Name);
                                page.Map[innerRow, innerCol].border = contentObjects.OfType<PdfCellBorderLayout>().FirstOrDefault(t => t.Name == m.Name);
                                page.Map[innerRow, innerCol].row = row + 1;
                                page.Map[innerRow, innerCol].col = col + 1;
                                innerCol++;
                                if (innerCol > mColCount)
                                {
                                    innerCol = col;
                                    innerRow++;
                                }
                            }
                        }
                        col = col+mColCount;
                        if (col >= colCount)
                        {
                            col = 0;
                            row++;
                        }
                    }
                    else if (child is PdfCellLayout l)
                    {
                        while (!string.IsNullOrEmpty(page.Map[row, col].Name))
                        {
                            col++;
                            if (col >= colCount)
                            {
                                col = 0;
                                row++;
                            }
                            if (row > rowCount)
                            {
                                break;
                            }
                        }
                        if (row < rowCount && col < colCount)
                        {
                            page.Map[row, col].Name = l.Name;
                            page.Map[row, col].cell = l;
                            page.Map[row, col].Type = PageMap.CellType.Normal;
                            page.Map[row, col].content = contentObjects.OfType<PdfCellContentLayout>().FirstOrDefault(t => t.Name == l.Name);
                            page.Map[row, col].border = contentObjects.OfType<PdfCellBorderLayout>().FirstOrDefault(t => t.Name == l.Name);
                            page.Map[row, col].row = row + 1;
                            page.Map[row, col].col = col + 1;
                        }
                        col++;
                        if (col >= colCount)
                        {
                            col = 0;
                            row++;
                        }
                    }
                }
                page.RemoveChild(page.ChildObjects[0]);
            }
        }

        private void ProocessPageAndCells(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfPagesLayout pages)
        {
            foreach (PdfPageLayout page in pages.ChildObjects)
            {
                var rowCount = page.ToRow - page.FromRow + 1;
                var colCount = page.ToCol - page.FromCol + 1;

                var activeVertical = new Dictionary<int, VerticalLineRun>();

                var activeHorizontal = new Dictionary<int, HorizontalLineRun>();
                Action<int> flushHorizontal = delegate (int row)
                {
                    HorizontalLineRun run;
                    if (activeHorizontal.TryGetValue(row, out run))
                    {
                        double y = page.Map[row, run.ColStart].cell.LocalPosition.Y;
                        double x1 = page.Map[row, run.ColStart].cell.LocalPosition.X;
                        double x2 = page.Map[row, run.ColEnd].cell.LocalPosition.X + page.Map[row, run.ColEnd].cell.Size.X;

                        page.GridLines.Add(new GridLine ( x1, y, x2, y ));
                        activeHorizontal.Remove(row);
                    }
                };
                Action<int, int> AddHorizontalSegment = delegate (int r, int c)
                {
                    HorizontalLineRun run;
                    if (activeHorizontal.TryGetValue(r, out run))
                    {
                        if (run.ColEnd == c - 1)
                        {
                            run.ColEnd = c;
                        }
                        else
                        {
                            flushHorizontal(r);
                            activeHorizontal[r] = new HorizontalLineRun { Row = r, ColStart = c, ColEnd = c };
                        }
                    }
                    else
                    {
                        activeHorizontal[r] = new HorizontalLineRun { Row = r, ColStart = c, ColEnd = c };
                    }
                };

                for (int row = 0; row < rowCount; row++)
                {
                    // flush horizontal runs not continued into this row
                    var hKeys = new List<int>(activeHorizontal.Keys);
                    foreach (int key in hKeys)
                        flushHorizontal(key);

                    for (int col = 0; col < colCount; col++)
                    {
                        var cell = page.Map[row, col];
                        PageMap leftCell = new PageMap();
                        PageMap rightCell = new PageMap();
                        if (col != 0) leftCell = page.Map[row, col - 1];
                        if (col != colCount - 1) rightCell = page.Map[row, col + 1];

                        //Check for text that spills into other cells. Used for gridlines

                        //Check text spill from left cell
                        if (leftCell.cell != null)
                        {
                            if (leftCell.content != null)
                            {
                                if (leftCell.content.RightTextSpillLength > 0d)
                                {
                                    cell.RightTextBucketSpill = leftCell.content.RightTextSpillLength - cell.cell.Size.X;
                                }
                            }
                            else if (leftCell.RightTextBucketSpill > 0d)
                            {
                                cell.RightTextBucketSpill = leftCell.RightTextBucketSpill - cell.cell.Size.X;
                            }
                        }
                        //check text spill to right //gör denna också för  left cell...
                        if (cell.content != null)
                        {
                            if (cell.content.LeftTextSpillLength > 0d)
                            {
                                double spill = cell.content.LeftTextSpillLength;
                                for (int i = col; i > 0; i--)
                                {
                                    var prevCell = page.Map[row, i - 1];
                                    if (prevCell.content != null)
                                    {
                                        cell.content.CreateClippingRect(cell.cell, prevCell.cell.LocalPosition.X + prevCell.cell.Size.X);
                                        break;
                                    }
                                    prevCell.LeftTextBucketSpill = spill - prevCell.cell.Size.X;
                                    spill = spill - prevCell.cell.Size.X;
                                    page.Map[row, i - 1] = prevCell;
                                    if (spill <= 0d) break;
                                }
                            }
                        }

                        //Collect gridlines

                        //Check Top edge
                        bool hasTop = row > 0;
                        var top = hasTop ? page.Map[row - 1, col] : default;

                        bool differentTop = !hasTop || top.cell != cell.cell;

                        bool borderTop = hasTop && (top.border != null &&
                            (top.border.BorderData.Bottom.BorderStyle != ExcelBorderStyle.None || cell.border.BorderData.Top.BorderStyle != ExcelBorderStyle.None));

                        if (differentTop && !borderTop)
                            AddHorizontalSegment(row, col);

                        //Check Bottom edge
                        if (row == page.RowCount - 1)
                        {
                            if (cell.border.BorderData.Bottom.BorderStyle != ExcelBorderStyle.None)
                            {
                                AddHorizontalSegment(row + 1, col);
                            }
                        }



                        if(page.Map[row, col].content != null) page.Map[row, col].content.CreateTextShape(dictionaries);
                        page.Map[row, col] = cell;
                    }
                }
                //create gridlines
                //make adjustments
                //AddHeaderFooter
                //remove unused cells
                //sort
            }
        }

        //Restore the positions of the content, move content children to page and remove content object.
        private void ConvertToPDFCoordiantes(PdfPageSettings pageSettings, PdfPagesLayout pages, ExcelWorksheet ws)
        {
            foreach (PdfPageLayout page in pages.ChildObjects)
            {
                page.ChildObjects[0].LocalPosition = new Vector2(pageSettings.Margins.LeftPu, pageSettings.Margins.TopPu);
                var contentObjects = page.ChildObjects[0].ChildObjects.ToList();
                foreach (var child in contentObjects)
                {
                    page.AddChild(child);
                    if (child is IShadingLayout iSl)
                        iSl.UpdateShadingPositionMatrix(pageSettings);
                    if (child is IBorderLayout iBl)
                        iBl.UpdateLocalBorderPosition();
                }
                page.RemoveChild(page.ChildObjects[0]);

                page.GenerateVerticalGridLines(ws);
                page.GenerateHorizontalGridLines(ws);
                page.GenerateBorderLines();
                //page.GenerateGridLines();
                page.ChildObjects.RemoveAll(x => x.Name.Contains("*")); //Remove all content with * in its name. Better approach would be to not add them at all, But they are needed for grid lines.
            }
        }

        //Make final adjustments and sort children for drawing order.
        private void AdjustAndSort(PdfPagesLayout pages, PdfDictionaries dictionaries)
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
                        contentLayout.CreateClippingRect(page.ChildObjects);//doe
                        contentLayout.CreateTextShape(dictionaries);
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
                    rightH =  ws.HeaderFooter.FirstHeader.RightAligned;
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
