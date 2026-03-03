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
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime;
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
            PopulatePages(pageSettings, dictionaries, WorksheetLayout, PagesLayout);
            //AddCellsToPageLayout(WorksheetLayout, PagesLayout);
            //HandleMergedCellsAndDrawings(pageSettings, dictionaries, WorksheetLayout, PagesLayout);
            MoveCellToPageFromContent(pageSettings, dictionaries, PagesLayout);
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

        private void PopulatePages(PdfPageSettings pageSettings, PdfDictionaries dictionaries, Transform WorksheetLayout, PdfPagesLayout pages)
        {
            var transforms = WorksheetLayout.ChildObjects.ToList();
            foreach (var t in transforms)
            {
                var cellBounds = t.GetGlobalBoundingbox();
                foreach (PdfPageLayout page in pages.ChildObjects)
                {
                    var contentbounds = page.ChildObjects[0].GetGlobalBoundingbox();
                    bool fullIntersect = IntersectsFully(contentbounds, cellBounds);
                    bool partialIntersect = Intersects(cellBounds, contentbounds);
                    if (t is PdfMergedCellLayout merged)
                    {
                        if (fullIntersect)
                        {
                            page.ChildObjects[0].AddChild(merged);
                            break;
                        }
                        else if (partialIntersect)
                        {
                            var copy = new PdfMergedCellLayout(dictionaries, merged.cell, merged.CellStyle, merged.LocalPosition.X, merged.LocalPosition.Y + merged.Size.Y, merged.Size.X, merged.Size.Y, merged.LocalScale.X, merged.LocalScale.Y, merged.LocalRotation, WorksheetLayout);
                            copy.Name = merged.Name;
                            copy.Z = merged.Z;
                            copy.address = merged.address;
                            page.ChildObjects[0].AddChild(copy);
                        }
                    }
                    else if (t is PdfCellLayout cell)
                    {
                        if (fullIntersect)
                        {
                            page.AddCell(cell);
                            break;
                        }
                    }
                    else if (t is PdfCellContentLayout content)
                    {
                        if (fullIntersect)
                        {
                            page.AddCell(content);
                            break;
                        }
                        else if (partialIntersect)
                        {
                            var copy = new PdfCellContentLayout(content.cell, content.CellStyle, pageSettings, content.LocalPosition.X, content.LocalPosition.Y, content.Size.X, content.Size.Y, content.LocalScale.X, content.LocalScale.Y, content.LocalRotation, WorksheetLayout, dictionaries);
                            copy.Name = content.Name;
                            copy.Z = content.Z;
                            page.ChildObjects[0].AddChild(copy);
                        }
                    }
                    else if (t is PdfCellBorderLayout border)
                    {
                        if (fullIntersect)
                        {
                            page.AddCell(border);
                            break;
                        }
                        else if (partialIntersect)
                        {
                            var copy = new PdfCellBorderLayout(border.cell, null, border.LocalPosition.X, border.LocalPosition.Y + border.Size.Y, border.Size.X, border.Size.Y, border.LocalScale.X, border.LocalScale.Y, border.LocalRotation, WorksheetLayout);
                            copy.Name = border.Name;
                            copy.Z = border.Z;
                            copy.BorderData = border.BorderData;
                            copy.range = border.range;
                            copy.IsMerged = border.IsMerged;
                            page.ChildObjects[0].AddChild(copy);
                        }
                    }
                    else if (t is PdfDrawingLayout drawing)
                    {
                        if (fullIntersect)
                        {
                            //
                        }
                        else if (partialIntersect)
                        {
                            var copy = new PdfDrawingLayout(null, drawing.LocalPosition.X, drawing.LocalPosition.Y, drawing.Size.X, drawing.Size.Y);
                            page.ChildObjects[0].AddChild(copy);
                        }
                    }
                }
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
                int index = 0;
                foreach (var page in pages.ChildObjects)
                {
                    if (Intersects(bounds, page.ChildObjects[0].GetGlobalBoundingbox()))
                    {
                        if (m is PdfMergedCellLayout)
                        {
                            var copy = new PdfMergedCellLayout(dictionaries, m.cell, m.CellStyle, m.LocalPosition.X, m.LocalPosition.Y + m.Size.Y, m.Size.X, m.Size.Y, m.LocalScale.X, m.LocalScale.Y, m.LocalRotation, WorksheetLayout);
                            copy.Name = m.Name;
                            copy.Z = m.Z;
                            copy.address = m.address;
                            page.ChildObjects[0].InsertChildAt(copy, index);
                        }
                        else if (c is PdfCellContentLayout)
                        {
                            var copy = new PdfCellContentLayout(c.cell, c.CellStyle, pageSettings, c.LocalPosition.X, c.LocalPosition.Y, c.Size.X, c.Size.Y, c.LocalScale.X, c.LocalScale.Y, c.LocalRotation, WorksheetLayout, dictionaries);
                            copy.Name = c.Name;
                            copy.Z = c.Z;
                            page.ChildObjects[0].InsertChildAt(copy, index);
                        }
                        else if (b is PdfCellBorderLayout)
                        {
                            var copy = new PdfCellBorderLayout(b.cell, null, b.LocalPosition.X, b.LocalPosition.Y + b.Size.Y, b.Size.X, b.Size.Y, b.LocalScale.X, b.LocalScale.Y, b.LocalRotation, WorksheetLayout);
                            copy.Name = b.Name;
                            copy.Z = b.Z;
                            copy.BorderData = b.BorderData;
                            copy.range = b.range;
                            copy.IsMerged = b.IsMerged;
                            page.ChildObjects[0].InsertChildAt(copy, index);
                        }
                        else if (d is PdfDrawingLayout) //NOT IMPLEMENTED
                        {
                            var copy = new PdfDrawingLayout(null, d.LocalPosition.X, d.LocalPosition.Y, d.Size.X, d.Size.Y);
                            page.ChildObjects[0].InsertChildAt(copy, index);
                        }
                    }
                    index++;
                }
            }
        }

        private void MoveCellToPageFromContent(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfPagesLayout pages)
        {
            foreach (PdfPageLayout page in pages.ChildObjects)
            {
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
                        var localFromRow = m.address._fromRow - page.FromRow;
                        var localFromCol = m.address._fromCol - page.FromCol;
                        var localToRow = m.address._toRow - page.FromRow;
                        var localToCol = m.address._toCol - page.FromCol;
                        int colCount = 0;
                        int rowCount = 0;

                        for (int r = localFromRow; r <= localToRow; r++)
                        {
                            if (r >= (page.ToRow - page.FromRow) + 1) break;
                            if (r < 0)
                            {
                                rowCount++;
                                continue;
                            }
                            for (int c = localFromCol; c <= localToCol; c++)
                            {
                                if (c >= (page.ToCol - page.FromCol) + 1) break;
                                if (c < 0)
                                {
                                    colCount++;
                                    continue;
                                }
                                page.Map[r, c].Name = ExcelRange.GetColumnLetter(m.cell._fromCol + colCount) + (m.cell._fromRow + rowCount);
                                page.Map[r, c].cell = m;
                                page.Map[r, c].Type = PageMap.CellType.Merged;
                                page.Map[r, c].content = contentObjects.OfType<PdfCellContentLayout>().FirstOrDefault(t => t.Name == m.Name);
                                page.Map[r, c].border = contentObjects.OfType<PdfCellBorderLayout>().FirstOrDefault(t => t.Name == m.Name);
                                page.Map[r, c].row = m.cell._fromRow + rowCount;
                                page.Map[r, c].col = m.cell._fromCol + colCount;
                                colCount++;
                            }
                            rowCount++;
                            colCount = 0;
                        }
                    }
                    else if (child is PdfCellLayout l)
                    {
                        var localFromRow = l.cell._fromRow - page.FromRow;
                        var localFromCol = l.cell._fromCol - page.FromCol;
                        page.Map[localFromRow, localFromCol].Name = l.Name;
                        page.Map[localFromRow, localFromCol].cell = l;
                        page.Map[localFromRow, localFromCol].Type = PageMap.CellType.Normal;
                        page.Map[localFromRow, localFromCol].content = contentObjects.OfType<PdfCellContentLayout>().FirstOrDefault(t => t.Name == l.Name);
                        page.Map[localFromRow, localFromCol].border = contentObjects.OfType<PdfCellBorderLayout>().FirstOrDefault(t => t.Name == l.Name);
                        page.Map[localFromRow, localFromCol].row = l.cell._fromRow;
                        page.Map[localFromRow, localFromCol].col = l.cell._fromCol;
                        page.Map[localFromRow, localFromCol].RightTextBucketSpill = page.Map[localFromRow, localFromCol].content != null ? page.Map[localFromRow, localFromCol].content.RightTextSpillLength : 0d;
                        page.Map[localFromRow, localFromCol].LeftTextBucketSpill = page.Map[localFromRow, localFromCol].content != null ? page.Map[localFromRow, localFromCol].content.LeftTextSpillLength : 0d;
                    }
                    else if (child is PdfCellContentLayout cl)
                    {
                        cl.CreateTextShape(dictionaries);
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
                Action<int> flushVertical = delegate (int col)
                {
                    VerticalLineRun run;
                    if (activeVertical.TryGetValue(col, out run))
                    {
                        double x = page.Map[run.RowStart, col].cell.LocalPosition.X;
                        double y1 = page.Map[run.RowStart, col].cell.LocalPosition.Y;
                        double y2 = page.Map[run.RowEnd, col].cell.LocalPosition.Y + page.Map[run.RowStart, col].cell.Size.Y;
                        page.GridLines.Add(new GridLine(x, y1, x, y2));
                        activeVertical.Remove(col);
                    }
                };
                Action<int, int> addVertical = delegate (int r, int c)
                {
                    VerticalLineRun run;
                    if (activeVertical.TryGetValue(c, out run))
                    {
                        if (run.RowEnd == r - 1)
                        {
                            run.RowEnd = r;
                        }
                        else
                        {
                            flushVertical(c);
                            activeVertical[c] = new VerticalLineRun { Col = c, RowStart = r, RowEnd = r };
                        }
                    }
                    else
                    {
                        activeVertical[c] = new VerticalLineRun { Col = c, RowStart = r, RowEnd = r };
                    }
                };

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

                    // flush vertical runs that didn't continue
                    var vKeys = new List<int>(activeVertical.Keys);
                    foreach (int key in vKeys)
                    {
                        VerticalLineRun run = activeVertical[key];
                        if (run.RowEnd < row)
                            flushVertical(key);
                    }

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

                        bool borderTop = hasTop && (top.border != null && cell.border != null &&
                            (top.border.BorderData.Bottom.BorderStyle != ExcelBorderStyle.None || cell.border.BorderData.Top.BorderStyle != ExcelBorderStyle.None));

                        if (differentTop && !borderTop)
                            AddHorizontalSegment(row, col);

                        //Check Bottom edge
                        if (row == page.RowCount - 1)
                        {
                            if (cell.border != null && cell.border.BorderData.Bottom.BorderStyle != ExcelBorderStyle.None)
                            {
                                AddHorizontalSegment(row + 1, col);
                            }
                        }

                        //Check Left edge
                        bool hasLeft = col > 0;
                        var left = hasLeft ? page.Map[row, col - 1] : default;

                        bool differentLeft = !hasLeft || left.cell != cell.cell;

                        bool spillLeft = hasLeft &&
                            (left.RightTextBucketSpill > 0 || cell.LeftTextBucketSpill > 0);

                        bool borderLeft = hasLeft && ((left.border != null && left.border.BorderData.Right.BorderStyle != ExcelBorderStyle.None) || (cell.border != null && cell.border.BorderData.Left.BorderStyle != ExcelBorderStyle.None));

                        if (differentLeft && !spillLeft && !borderLeft)
                            addVertical(row, col);

                        //Check Right edge
                        if (col == page.ColCount - 1)
                        {
                            if (cell.border != null && cell.border.BorderData.Right.BorderStyle != ExcelBorderStyle.None &&
                                cell.RightTextBucketSpill <= 0)
                            {
                                //addVertical(row, col + 1);
                            }
                        }


                        /*if (page.Map[row, col].content != null) page.Map[row, col].content.CreateTextShape(dictionaries);*/
                        page.Map[row, col] = cell;
                    }
                }
                // final flush
                foreach (int key in new List<int>(activeVertical.Keys))
                    flushVertical(key);

                foreach (int key in new List<int>(activeHorizontal.Keys))
                    flushHorizontal(key);
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
