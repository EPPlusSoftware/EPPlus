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
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.TextShaping;
using EPPlus.Graphics;
using EPPlus.Graphics.Math;
using EPPlus.Graphics.Units;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.HeaderFooterTextFormat;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
            Stopwatch sw = Stopwatch.StartNew();

            Name = worksheet.Name + " Catalog";
            var WorksheetLayout = AddChild(new PdfWorksheetLayout(worksheet, pageSettings, dictionaries));
            sw.Stop();
            var t1 = sw.ElapsedMilliseconds;
            sw.Reset();
            sw.Start();

            //CreateFontSubsets(pageSettings, dictionaries.Fonts);
            var PagesLayout = CreatePagesLayoutObject();
            sw.Stop();
            var t2 = sw.ElapsedMilliseconds;
            sw.Reset();
            sw.Start();

            CreatePageLayoutObjects(worksheet, pageSettings, WorksheetLayout as PdfWorksheetLayout, PagesLayout);
            sw.Stop();
            var t3 = sw.ElapsedMilliseconds;
            sw.Reset();
            sw.Start();

            AddHeaderFooter(worksheet, pageSettings, dictionaries, PagesLayout);
            sw.Stop();
            var t8 = sw.ElapsedMilliseconds;
            sw.Reset();
            sw.Start();

            PopulatePages(pageSettings, dictionaries, WorksheetLayout, PagesLayout);
            sw.Stop();
            var t4 = sw.ElapsedMilliseconds;
            sw.Reset();
            sw.Start();

            //AddCellsToPageLayout(WorksheetLayout, PagesLayout);
            //HandleMergedCellsAndDrawings(pageSettings, dictionaries, WorksheetLayout, PagesLayout);
            MoveCellToPageFromContent(pageSettings, dictionaries, PagesLayout);
            sw.Stop();
            var t5 = sw.ElapsedMilliseconds;
            sw.Reset();
            sw.Start();

            ProocessPageAndCells(pageSettings, dictionaries, PagesLayout);
            sw.Stop();
            var t6 = sw.ElapsedMilliseconds;
            sw.Reset();
            sw.Start();

            //ConvertToPDFCoordiantes(pageSettings, PagesLayout, worksheet);
            //AdjustAndSort(PagesLayout, dictionaries);
            RemoveChild(WorksheetLayout);
            sw.Stop();
            var t7 = sw.ElapsedMilliseconds;
            sw.Reset();
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

        //Internal class for storing bounds for pages
        private class PageData
        {
            public PdfPageLayout Page;
            public Rect Bounds; // replace BoundingBox with whatever your bounds type is

            public PageData(PdfPageLayout page, Rect bounds)
            {
                Page = page;
                Bounds = bounds;
            }
        }

        private Dictionary<IFontProvider, TextShaper> shaperCache = new Dictionary<IFontProvider, TextShaper>();
        private Dictionary<IFontProvider, TextLayoutEngine> layoutEngineCache = new Dictionary<IFontProvider, TextLayoutEngine>();

        //Move cells to their overlapping pages.
        private void PopulatePages(PdfPageSettings pageSettings, PdfDictionaries dictionaries, Transform WorksheetLayout, PdfPagesLayout pages)
        {

            var pageData = new List<PageData>();
            foreach (PdfPageLayout p in pages.ChildObjects)
            {
                pageData.Add(new PageData(p, p.ChildObjects[0].GetGlobalBoundingbox()));

                var hfs = p.ChildObjects.Where(x => x is PdfHeaderFooterLayout).ToArray();
                foreach(var hf in hfs )
                    LayoutAndShapeText(pageSettings, dictionaries, shaperCache, layoutEngineCache, (PdfHeaderFooterLayout)hf);
            }
            pageData.Sort((a, b) => a.Bounds.Top.CompareTo(b.Bounds.Top));

            var transforms = new List<Transform>(WorksheetLayout.ChildObjects);
            foreach (var t in transforms)
            {
                var cellBounds = t.GetGlobalBoundingbox();

                // Pass 1: check if ANY page fully contains this transform. 
                // If so, we use only that page and skip all partial intersects.
                PageData fullIntersectPage = null;
                foreach (var pd in pageData)
                {
                    if (pd.Bounds.Top > cellBounds.Bottom) break;
                    if (pd.Bounds.Bottom < cellBounds.Top) continue;

                    if (IntersectsFully(pd.Bounds, cellBounds))
                    {
                        fullIntersectPage = pd;
                        break;
                    }
                }
                // Pass 2: assign to pages.
                foreach (var pd in pageData)
                {
                    if (pd.Bounds.Top > cellBounds.Bottom) break;
                    if (pd.Bounds.Bottom < cellBounds.Top) continue;

                    bool fullIntersect = fullIntersectPage != null && pd == fullIntersectPage;
                    bool partialIntersect = fullIntersectPage == null && !fullIntersect && Intersects(cellBounds, pd.Bounds);

                    if (!fullIntersect && !partialIntersect) continue;

                    var page = pd.Page;

                    if (t is PdfMergedCellLayout merged)
                    {
                        if (fullIntersect)
                        {
                            page.ChildObjects[0].AddChild(merged);
                            break;
                        }
                        else
                        {
                            var copy = new PdfMergedCellLayout(dictionaries, merged.cell, merged.CellStyle,
                                merged.LocalPosition.X, merged.LocalPosition.Y + merged.Size.Y,
                                merged.Size.X, merged.Size.Y,
                                merged.LocalScale.X, merged.LocalScale.Y,
                                merged.LocalRotation, WorksheetLayout);
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
                    else if (t is PdfCellContentLayout cellContent)
                    {
                        if (fullIntersect)
                        {
                            page.AddCell(cellContent);
                            LayoutAndShapeText(pageSettings, dictionaries, shaperCache, layoutEngineCache, cellContent);
                            break;
                        }
                        else
                        {
                            var copy = new PdfCellContentLayout(cellContent.cell, cellContent.CellStyle, pageSettings,
                                cellContent.LocalPosition.X, cellContent.LocalPosition.Y,
                                cellContent.Size.X, cellContent.Size.Y,
                                cellContent.LocalScale.X, cellContent.LocalScale.Y,
                                cellContent.LocalRotation, WorksheetLayout, dictionaries);
                            copy.Name = cellContent.Name;
                            copy.Z = cellContent.Z;
                            page.ChildObjects[0].AddChild(copy);
                            LayoutAndShapeText(pageSettings, dictionaries, shaperCache, layoutEngineCache, copy);
                        }
                    }
                    else if (t is PdfCellBorderLayout border)
                    {
                        if (fullIntersect)
                        {
                            page.AddCell(border);
                            break;
                        }
                        else
                        {
                            var copy = new PdfCellBorderLayout(border.cell, null,
                                border.LocalPosition.X, border.LocalPosition.Y + border.Size.Y,
                                border.Size.X, border.Size.Y,
                                border.LocalScale.X, border.LocalScale.Y,
                                border.LocalRotation, WorksheetLayout);
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
                            break;
                        else
                        {
                            var copy = new PdfDrawingLayout(null,
                                drawing.LocalPosition.X, drawing.LocalPosition.Y,
                                drawing.Size.X, drawing.Size.Y);
                            page.ChildObjects[0].AddChild(copy);
                        }
                    }
                }
            }
        }

        //Font handling, Text shaping and layouting wrapped text
        private static void LayoutAndShapeText(PdfPageSettings pageSettings, PdfDictionaries dictionaries, Dictionary<IFontProvider, TextShaper> shaperCache, Dictionary<IFontProvider, TextLayoutEngine> layoutEngineCache, ITextLayout text)
        {
            for (int i = 0; i < text.TextFormats.Count; i++)
            {
                var fd = text.TextFormats[i];
                fd.FontProvider = dictionaries.Fonts[fd.FullFontName].fontSubsetManager.CreateSubsettedProvider();
                //Wraptext TODO


                if (!shaperCache.TryGetValue(fd.FontProvider, out var shaper))
                {
                    shaper = new TextShaper(fd.FontProvider);
                    shaperCache[fd.FontProvider] = shaper;
                }

                if (!layoutEngineCache.TryGetValue(fd.FontProvider, out var layoutEngine))
                {
                    layoutEngine = new TextLayoutEngine(shaper);
                    layoutEngineCache[fd.FontProvider] = layoutEngine;
                }

                var options = ShapingOptions.Default;
                options.ApplyPositioning = true;
                options.ApplySubstitutions = true;

                // Shape the text
                var shaped = shaper.Shape(fd.Text, options);

                // Get all fonts used in this paragraph (primary + fallbacks like emoji)
                var usedFonts = shaper.GetUsedFonts().ToList();
                var fontIdMap = new Dictionary<byte, string>();

                // Register each font globally and map FontId → Resource Index
                for (byte fontId = 0; fontId < usedFonts.Count; fontId++)
                {
                    var font = usedFonts[fontId];

                    // Add to global tracking if new
                    if (!dictionaries.Fonts.ContainsKey(font.FullName))
                    {
                        int label = 1;
                        if (dictionaries.Fonts.Count > 0)
                        {
                            label = dictionaries.Fonts.Last().Value.labelNumber + 1;
                        }
                        dictionaries.Fonts.Add(font.FullName, new PdfFontResource(font.FullName, font.NameTable.GetSubfamilyEnum(), label, pageSettings));
                    }
                    fontIdMap[fontId] = dictionaries.Fonts[font.FullName].Label;
                }
                text.TextLayoutEngine = layoutEngine;
                fd.ShapedText = shaped;
                fd.FontIdMap = fontIdMap;
                fd.UsedFonts = usedFonts;
                text.TextFormats[i] = fd;
                shaper.ResetFontTracking();
            }
        }

        //Create a map for page
        private void MoveCellToPageFromContent(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfPagesLayout pages)
        {
            List<string> entriesToRemove = new();
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
                        bool remove = true;

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
                                remove = false;
                            }
                            rowCount++;
                            colCount = 0;
                        }
                        if (remove)
                        {
                            entriesToRemove.Add(child.Name);
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
                }
                page.RemoveChild(page.ChildObjects[0]);
                foreach (var entry in entriesToRemove)
                {
                    var cell = contentObjects.OfType<PdfMergedCellLayout>().FirstOrDefault(t => t.Name == entry);
                    var content = contentObjects.OfType<PdfCellContentLayout>().FirstOrDefault(t => t.Name == entry);
                    var border = contentObjects.OfType<PdfCellBorderLayout>().FirstOrDefault(t => t.Name == entry);
                    page.RemoveChild(content);
                    page.RemoveChild(border);
                    page.RemoveChild(cell);
                }
            }
        }

        private void ProocessPageAndCells(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfPagesLayout pages)
        {
            foreach (PdfPageLayout page in pages.ChildObjects)
            {
                var rowCount = page.ToRow - page.FromRow + 1;
                var colCount = page.ToCol - page.FromCol + 1;

                // ── Interior vertical lines ──────────────────────────────────────────
                var activeVertical = new Dictionary<int, VerticalLineRun>();
                Action<int> flushVertical = delegate (int col)
                {
                    VerticalLineRun run;
                    if (activeVertical.TryGetValue(col, out run))
                    {
                        double x = page.Map[run.RowStart, col].cell.LocalPosition.X;
                        double y1 = page.Map[run.RowStart, col].cell.LocalPosition.Y;
                        double y2 = page.Map[run.RowEnd, col].cell.LocalPosition.Y
                                  + page.Map[run.RowEnd, col].cell.Size.Y;
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
                            run.RowEnd = r;
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

                // ── Interior horizontal lines ────────────────────────────────────────
                var activeHorizontal = new Dictionary<int, HorizontalLineRun>();
                Action<int> flushHorizontal = delegate (int row)
                {
                    HorizontalLineRun run;
                    if (activeHorizontal.TryGetValue(row, out run))
                    {
                        double y = page.Map[row, run.ColStart].cell.LocalPosition.Y;
                        double x1 = page.Map[row, run.ColStart].cell.LocalPosition.X;
                        double x2 = page.Map[row, run.ColEnd].cell.LocalPosition.X
                                  + page.Map[row, run.ColEnd].cell.Size.X;
                        page.GridLines.Add(new GridLine(x1, y, x2, y));
                        activeHorizontal.Remove(row);
                    }
                };
                Action<int, int> AddHorizontalSegment = delegate (int r, int c)
                {
                    HorizontalLineRun run;
                    if (activeHorizontal.TryGetValue(r, out run))
                    {
                        if (run.ColEnd == c - 1)
                            run.ColEnd = c;
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

                // ── Main loop ────────────────────────────────────────────────────────
                for (int row = 0; row < rowCount; row++)
                {
                    // Flush horizontal runs not continued into this row
                    foreach (int key in new List<int>(activeHorizontal.Keys))
                        flushHorizontal(key);

                    // Flush vertical runs that didn't continue
                    foreach (int key in new List<int>(activeVertical.Keys))
                    {
                        if (activeVertical[key].RowEnd < row)
                            flushVertical(key);
                    }

                    for (int col = 0; col < colCount; col++)
                    {
                        var cell = page.Map[row, col];
                        PageMap leftCell = new PageMap();
                        PageMap rightCell = new PageMap();
                        if (col != 0) leftCell = page.Map[row, col - 1];
                        if (col != colCount - 1) rightCell = page.Map[row, col + 1];

                        // ── Text-spill propagation (unchanged) ───────────────────────
                        if (leftCell.cell != null)
                        {
                            if (leftCell.content != null)
                            {
                                if (leftCell.content.RightTextSpillLength > 0d)
                                    cell.RightTextBucketSpill = leftCell.content.RightTextSpillLength - cell.cell.Size.X;
                            }
                            else if (leftCell.RightTextBucketSpill > 0d)
                            {
                                cell.RightTextBucketSpill = leftCell.RightTextBucketSpill - cell.cell.Size.X;
                            }
                        }

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
                                    spill -= prevCell.cell.Size.X;
                                    page.Map[row, i - 1] = prevCell;
                                    if (spill <= 0d) break;
                                }
                            }

                            cell.content.GidsAndCharMap(dictionaries);
                        }

                        // ── Gridline collection ──────────────────────────────────────

                        // Interior horizontal — only between two rows (skip top page edge)
                        if (row > 0)
                        {
                            var t = page.Map[row - 1, col];
                            bool differentTop = t.cell != cell.cell;
                            bool borderTop = (t.border != null && t.border.BorderData.Bottom.BorderStyle != ExcelBorderStyle.None) ||
                                       (cell.border != null && cell.border.BorderData.Top.BorderStyle != ExcelBorderStyle.None);

                            if (differentTop && !borderTop)
                                AddHorizontalSegment(row - 1, col);
                        }

                        // Interior vertical — only between two columns (skip left page edge)
                        if (col > 0)
                        {
                            var l = page.Map[row, col - 1];
                            bool differentLeft = l.cell != cell.cell;
                            bool spillLeft = l.RightTextBucketSpill > 0 || cell.LeftTextBucketSpill > 0;
                            bool borderLeft = (l.border != null && l.border.BorderData.Right.BorderStyle != ExcelBorderStyle.None) ||
                                                (cell.border != null && cell.border.BorderData.Left.BorderStyle != ExcelBorderStyle.None);

                            if (differentLeft && !spillLeft && !borderLeft)
                                addVertical(row, col);
                        }

                        page.Map[row, col] = cell;
                    }
                }

                // ── Final flush ──────────────────────────────────────────────────────
                foreach (int key in new List<int>(activeVertical.Keys))
                    flushVertical(key);
                foreach (int key in new List<int>(activeHorizontal.Keys))
                    flushHorizontal(key);

                var tl = page.Map[0, 0].cell; // bottom-left
                var br = page.Map[0, colCount - 1].cell; // bottom-right
                var bl = page.Map[rowCount - 1, 0].cell; // top-left
                var tr = page.Map[rowCount - 1, colCount - 1].cell; // top-right

                double left = bl.LocalPosition.X;
                double right = br.LocalPosition.X + br.Size.X;
                double bottom = bl.LocalPosition.Y;
                double top = tl.LocalPosition.Y + tl.Size.Y;

                page.BorderLines.Add(new GridLine(left, bottom, right, bottom)); // bottom
                page.BorderLines.Add(new GridLine(left, top, right, top));    // top
                page.BorderLines.Add(new GridLine(left, bottom, left, top));    // left
                page.BorderLines.Add(new GridLine(right, bottom, right, top));    // right
            }
        }

        //Create gridlines
        //private void ProocessPageAndCells(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfPagesLayout pages)
        //{
        //    foreach (PdfPageLayout page in pages.ChildObjects)
        //    {
        //        var rowCount = page.ToRow - page.FromRow + 1;
        //        var colCount = page.ToCol - page.FromCol + 1;

        //        var activeVertical = new Dictionary<int, VerticalLineRun>();
        //        Action<int> flushVertical = delegate (int col)
        //        {
        //            VerticalLineRun run;
        //            if (activeVertical.TryGetValue(col, out run))
        //            {
        //                double x = page.Map[run.RowStart, col].cell.LocalPosition.X;
        //                double y1 = page.Map[run.RowStart, col].cell.LocalPosition.Y;
        //                double y2 = page.Map[run.RowEnd, col].cell.LocalPosition.Y + page.Map[run.RowStart, col].cell.Size.Y;
        //                page.GridLines.Add(new GridLine(x, y1, x, y2));
        //                activeVertical.Remove(col);
        //            }
        //        };
        //        Action<int, int> addVertical = delegate (int r, int c)
        //        {
        //            VerticalLineRun run;
        //            if (activeVertical.TryGetValue(c, out run))
        //            {
        //                if (run.RowEnd == r - 1)
        //                {
        //                    run.RowEnd = r;
        //                }
        //                else
        //                {
        //                    flushVertical(c);
        //                    activeVertical[c] = new VerticalLineRun { Col = c, RowStart = r, RowEnd = r };
        //                }
        //            }
        //            else
        //            {
        //                activeVertical[c] = new VerticalLineRun { Col = c, RowStart = r, RowEnd = r };
        //            }
        //        };

        //        var activeHorizontal = new Dictionary<int, HorizontalLineRun>();
        //        Action<int> flushHorizontal = delegate (int row)
        //        {
        //            HorizontalLineRun run;
        //            if (activeHorizontal.TryGetValue(row, out run))
        //            {
        //                double y = page.Map[row, run.ColStart].cell.LocalPosition.Y;
        //                double x1 = page.Map[row, run.ColStart].cell.LocalPosition.X;
        //                double x2 = page.Map[row, run.ColEnd].cell.LocalPosition.X + page.Map[row, run.ColEnd].cell.Size.X;
        //                page.GridLines.Add(new GridLine(x1, y, x2, y));
        //                activeHorizontal.Remove(row);
        //            }
        //        };
        //        Action<int, int> AddHorizontalSegment = delegate (int r, int c)
        //        {
        //            HorizontalLineRun run;
        //            if (activeHorizontal.TryGetValue(r, out run))
        //            {
        //                if (run.ColEnd == c - 1)
        //                {
        //                    run.ColEnd = c;
        //                }
        //                else
        //                {
        //                    flushHorizontal(r);
        //                    activeHorizontal[r] = new HorizontalLineRun { Row = r, ColStart = c, ColEnd = c };
        //                }
        //            }
        //            else
        //            {
        //                activeHorizontal[r] = new HorizontalLineRun { Row = r, ColStart = c, ColEnd = c };
        //            }
        //        };

        //        for (int row = 0; row < rowCount; row++)
        //        {
        //            // flush horizontal runs not continued into this row
        //            var hKeys = new List<int>(activeHorizontal.Keys);
        //            foreach (int key in hKeys)
        //                flushHorizontal(key);

        //            // flush vertical runs that didn't continue
        //            var vKeys = new List<int>(activeVertical.Keys);
        //            foreach (int key in vKeys)
        //            {
        //                VerticalLineRun run = activeVertical[key];
        //                if (run.RowEnd < row)
        //                    flushVertical(key);
        //            }

        //            for (int col = 0; col < colCount; col++)
        //            {
        //                var cell = page.Map[row, col];
        //                PageMap leftCell = new PageMap();
        //                PageMap rightCell = new PageMap();
        //                if (col != 0) leftCell = page.Map[row, col - 1];
        //                if (col != colCount - 1) rightCell = page.Map[row, col + 1];

        //                //Check for text that spills into other cells. Used for gridlines

        //                //Check text spill from left cell
        //                if (leftCell.cell != null)
        //                {
        //                    if (leftCell.content != null)
        //                    {
        //                        if (leftCell.content.RightTextSpillLength > 0d)
        //                        {
        //                            cell.RightTextBucketSpill = leftCell.content.RightTextSpillLength - cell.cell.Size.X;
        //                        }
        //                    }
        //                    else if (leftCell.RightTextBucketSpill > 0d)
        //                    {
        //                        cell.RightTextBucketSpill = leftCell.RightTextBucketSpill - cell.cell.Size.X;
        //                    }
        //                }
        //                //check text spill to right //gör denna också för  left cell...
        //                if (cell.content != null)
        //                {
        //                    if (cell.content.LeftTextSpillLength > 0d)
        //                    {
        //                        double spill = cell.content.LeftTextSpillLength;
        //                        for (int i = col; i > 0; i--)
        //                        {
        //                            var prevCell = page.Map[row, i - 1];
        //                            if (prevCell.content != null)
        //                            {
        //                                cell.content.CreateClippingRect(cell.cell, prevCell.cell.LocalPosition.X + prevCell.cell.Size.X);
        //                                break;
        //                            }
        //                            prevCell.LeftTextBucketSpill = spill - prevCell.cell.Size.X;
        //                            spill = spill - prevCell.cell.Size.X;
        //                            page.Map[row, i - 1] = prevCell;
        //                            if (spill <= 0d) break;
        //                        }
        //                    }

        //                    cell.content.GidsAndCharMap(dictionaries);
        //                }

        //                //Collect gridlines

        //                //Check Top edge
        //                bool hasTop = row > 0;
        //                var top = hasTop ? page.Map[row - 1, col] : default;

        //                bool differentTop = !hasTop || top.cell != cell.cell;

        //                bool borderTop = hasTop && (top.border != null && cell.border != null &&
        //                    (top.border.BorderData.Bottom.BorderStyle != ExcelBorderStyle.None || cell.border.BorderData.Top.BorderStyle != ExcelBorderStyle.None));

        //                if (differentTop && !borderTop)
        //                    AddHorizontalSegment(row, col);

        //                //Check Bottom edge
        //                if (row == page.RowCount - 1)
        //                {
        //                    if (cell.border != null && cell.border.BorderData.Bottom.BorderStyle != ExcelBorderStyle.None)
        //                    {
        //                        AddHorizontalSegment(row + 1, col);
        //                    }
        //                }

        //                //Check Left edge
        //                bool hasLeft = col > 0;
        //                var left = hasLeft ? page.Map[row, col - 1] : default;

        //                bool differentLeft = !hasLeft || left.cell != cell.cell;

        //                bool spillLeft = hasLeft &&
        //                    (left.RightTextBucketSpill > 0 || cell.LeftTextBucketSpill > 0);

        //                bool borderLeft = hasLeft && ((left.border != null && left.border.BorderData.Right.BorderStyle != ExcelBorderStyle.None) || (cell.border != null && cell.border.BorderData.Left.BorderStyle != ExcelBorderStyle.None));

        //                if (differentLeft && !spillLeft && !borderLeft)
        //                    addVertical(row, col);

        //                //Check Right edge
        //                if (col == page.ColCount - 1)
        //                {
        //                    if (cell.border != null && cell.border.BorderData.Right.BorderStyle != ExcelBorderStyle.None &&
        //                        cell.RightTextBucketSpill <= 0)
        //                    {
        //                        //addVertical(row, col + 1);
        //                    }
        //                }


        //                /*if (page.Map[row, col].content != null) page.Map[row, col].content.CreateTextShape(dictionaries);*/
        //                page.Map[row, col] = cell;
        //            }
        //        }
        //        // final flush
        //        foreach (int key in new List<int>(activeVertical.Keys))
        //            flushVertical(key);

        //        foreach (int key in new List<int>(activeHorizontal.Keys))
        //            flushHorizontal(key);
        //        //create gridlines
        //        //make adjustments
        //        //AddHeaderFooter
        //        //remove unused cells
        //        //sort
        //    }
        //}

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
                        //contentLayout.CreateTextShape(dictionaries);
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

        private void AddHeaderFooter(ExcelWorksheet ws, PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfPagesLayout pages)
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
                var lh = (PdfHeaderFooterLayout)pages.ChildObjects[i].AddChild(new PdfHeaderFooterLayout(leftH, ws, pageSettings, dictionaries, pageNumber, pages.ChildObjects.Count));
                lh.LocalPosition = new Vector2(pageSettings.Margins.LeftPu, pageSettings.PageSize.HeightPu - pageSettings.Margins.HeaderPu);
                lh.AdjustPositionByTextLength('l', 'h');
                var ch = (PdfHeaderFooterLayout)pages.ChildObjects[i].AddChild(new PdfHeaderFooterLayout(centerH, ws, pageSettings, dictionaries, pageNumber, pages.ChildObjects.Count));
                ch.LocalPosition = new Vector2(pageSettings.PageSize.WidthPu / 2d, pageSettings.PageSize.HeightPu - pageSettings.Margins.HeaderPu);
                ch.AdjustPositionByTextLength('c', 'h');
                var rh = (PdfHeaderFooterLayout)pages.ChildObjects[i].AddChild(new PdfHeaderFooterLayout(rightH, ws, pageSettings, dictionaries, pageNumber, pages.ChildObjects.Count));
                rh.LocalPosition = new Vector2(pageSettings.PageSize.WidthPu - pageSettings.Margins.RightPu, pageSettings.PageSize.HeightPu - pageSettings.Margins.HeaderPu);
                rh.AdjustPositionByTextLength('r', 'h');
                var lf = (PdfHeaderFooterLayout)pages.ChildObjects[i].AddChild(new PdfHeaderFooterLayout(leftF, ws, pageSettings, dictionaries, pageNumber, pages.ChildObjects.Count));
                lf.LocalPosition = new Vector2(pageSettings.Margins.LeftPu, pageSettings.Margins.FooterPu);
                var cf = (PdfHeaderFooterLayout)pages.ChildObjects[i].AddChild(new PdfHeaderFooterLayout(centerF, ws, pageSettings, dictionaries, pageNumber, pages.ChildObjects.Count));
                cf.LocalPosition = new Vector2(pageSettings.PageSize.WidthPu / 2d, pageSettings.Margins.FooterPu);
                cf.AdjustPositionByTextLength('c', 'f');
                var rf = (PdfHeaderFooterLayout)pages.ChildObjects[i].AddChild(new PdfHeaderFooterLayout(rightF, ws, pageSettings, dictionaries, pageNumber, pages.ChildObjects.Count));
                rf.LocalPosition = new Vector2(pageSettings.PageSize.WidthPu - pageSettings.Margins.RightPu, pageSettings.Margins.FooterPu);
                rf.AdjustPositionByTextLength('r', 'f');
                pageNumber++;
            }
        }
    }
}
