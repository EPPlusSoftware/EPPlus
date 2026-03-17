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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.HeaderFooterTextFormat;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security;

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
            var s0 = this.ToHierarchyString();

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
            var s1 = this.ToHierarchyString();

            //AddCellsToPageLayout(WorksheetLayout, PagesLayout);
            //HandleMergedCellsAndDrawings(pageSettings, dictionaries, WorksheetLayout, PagesLayout);
            MoveCellToPageFromContent(pageSettings, dictionaries, PagesLayout);
            sw.Stop();
            var t5 = sw.ElapsedMilliseconds;
            sw.Reset();
            sw.Start();
            var s2 = this.ToHierarchyString();

            ProocessPageAndCells(pageSettings, dictionaries, PagesLayout);
            sw.Stop();
            var t6 = sw.ElapsedMilliseconds;
            sw.Reset();
            sw.Start();

            //ConvertToPDFCoordiantes(pageSettings, PagesLayout, worksheet);
            AdjustAndSort(PagesLayout, dictionaries);
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

            /* when calculating pages we need to take row and column headings width and height into consideration we could start our for loops and 0 and check if j==0 and add row width and then just proceed like usual and
             * every time we hit a break we then add row width again. these are found in worksheet layout to use where we calculate them. 
             *
             * Next step in populatePages? would be to add additional cells for these headings when we have a new page first row is creating new PdfCellLayout and PdfCellContentLayout. This might conflict with text shaping since we now add new text while also shaping text.
             * A solution to this could be when doing the worksheet layout is to check dimensions and add A-Z, 1-9 as needed to a new font entry in dictionaries.Fonts
             */

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
                if (worksheet.Column(j).PageBreak)
                {
                    xBreaks.Add(currentWidth);
                    boundsWidth = currentWidth + pageSettings.ContentBounds.Width;
                }
            }
                xBreaks.Add(currentWidth);
            //Get y cooridiantes to break for new page
            List<double> yBreaks = new List<double>() { 0d };
            double currentHeight = 0, boundsHeight = -pageSettings.ContentBounds.Height;
            for (int i = 1; i <= worksheet.Dimension._toRow; i++)
            {
                if (worksheet.Row(i).Hidden) { continue; }
                var height = UnitConversion.ExcelRowHeightToPoints(worksheet.Row(i).Height);
                if (currentHeight - height <= boundsHeight)
                {
                    yBreaks.Add(currentHeight);
                    boundsHeight = currentHeight - pageSettings.ContentBounds.Height;
                }
                currentHeight -= height;
                if (worksheet.Row(i).PageBreak)
                {
                    yBreaks.Add(currentHeight);
                    boundsHeight = currentHeight - pageSettings.ContentBounds.Height;
                }
            }
                yBreaks.Add(currentHeight);
            //calculate number of pages needed based on contentBounds and worksheetLayout.Size
            int horizontalPageCount = xBreaks.Count-1; //System.Math.Max(1, (int)System.Math.Ceiling(worksheetLayout.Size.X / pageSettings.ContentBounds.Width));
            int verticalPageCount = yBreaks.Count-1; //System.Math.Max(1, (int)System.Math.Ceiling(System.Math.Abs( worksheetLayout.Size.Y) / pageSettings.ContentBounds.Height));
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
                //ContentLayout
                double width = xBreaks[col + 1] - xBreaks[col];
                double height = Math.Abs( yBreaks[row] - yBreaks[row + 1]);
                PdfContentLayout content = new PdfContentLayout(0, 0, width, height);
                content.Name = "Content " + (i + 1);
                page.AddChild(content);
                content.Position = new Vector2(xBreaks[col], yBreaks[row+1]);
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
        //Font handling, Text shaping and layouting wrapped text
        private static void LayoutAndShapeText(PdfPageSettings pageSettings, PdfDictionaries dictionaries, Dictionary<IFontProvider, TextShaper> shaperCache, Dictionary<IFontProvider, TextLayoutEngine> layoutEngineCache, ITextLayout text)
        {
            var totalTextLength = 0d;
            var maxLineHeight = 0d;
            for (int i = 0; i < text.TextFormats.Count; i++)
            {
                var fd = text.TextFormats[i];
                fd.FontProvider = dictionaries.Fonts[fd.FullFontName].fontSubsetManager.CreateSubsettedProvider();

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

                var shaped = shaper.Shape(fd.Text, options);
                var usedFonts = shaper.GetUsedFonts().ToList();
                var fontIdMap = new Dictionary<byte, string>();

                var allProviderFonts = fd.FontProvider.GetAllFonts().ToList();

                for (byte fontId = 0; fontId < usedFonts.Count; fontId++)
                {
                    var font = usedFonts[fontId];

                    if (!dictionaries.Fonts.ContainsKey(font.FullName))
                    {
                        int label = 1;
                        if (dictionaries.Fonts.Count > 0)
                        {
                            label = dictionaries.Fonts.Last().Value.labelNumber + 1;
                        }
                        var fontResource = new PdfFontResource(font.FullName, font.NameTable.GetSubfamilyEnum(), label, pageSettings);
                        fontResource.fontData = font;
                        dictionaries.Fonts.Add(font.FullName, fontResource);
                    }
                    fontIdMap[fontId] = dictionaries.Fonts[font.FullName].Label;
                }

                text.TextLayoutEngine = layoutEngine;
                fd.ShapedText = shaped;
                var textWdith = fd.ShapedText.GetWidthInPoints((float)fd.FontSize);
                var textHeight = fd.ShapedText.GetLineHeightInPoints((float)fd.FontSize);
                fd.TextLength = textWdith;
                fd.TextHeight = textHeight;
                totalTextLength += textWdith;
                maxLineHeight = Math.Max(textHeight, maxLineHeight);
                fd.FontIdMap = fontIdMap;
                fd.UsedFonts = usedFonts;
                text.TextFormats[i] = fd;
                shaper.ResetFontTracking();
            }
            //saker här sen
            if (text is PdfCellContentLayout ccl)
            {
                ccl.TextLength = totalTextLength;
                ccl.TextHeight = maxLineHeight;
                ccl.CalculateTextSpill(ccl.Size.X, ccl.CellAlignmentData.TextRotation);
                ccl.LocalPosition = ccl.CalculateAlignmentPositionAndTextOffsets(ccl.cell, ccl.LocalPosition.X, ccl.LocalPosition.Y, ccl.Size.X, ccl.Size.Y);
                ccl.CheckClipping(ccl.cell, ccl.LocalPosition.X, ccl.LocalPosition.Y, ccl.Size.X, ccl.Size.Y);
            }
        }

        //Create a map for page
        private void MoveCellToPageFromContent(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfPagesLayout pages)
        {
            List<string> entriesToRemove = new();
            foreach (PdfPageLayout page in pages.ChildObjects)
            {
                page.CreateMap();
                var x = pageSettings.Margins.LeftPu;
                if (pageSettings.CenterOnPageHorizontally)
                {
                    var w = pageSettings.PageSize.WidthPu - pageSettings.Margins.LeftPu - pageSettings.Margins.RightPu;
                    x = pageSettings.Margins.LeftPu + (w - page.ChildObjects[0].Size.X) / 2;
                }
                var y = pageSettings.PageSize.HeightPu - pageSettings.Margins.TopPu - page.ChildObjects[0].Size.Y;
                if (pageSettings.CenterOnPageVertically)
                {
                    var h = pageSettings.PageSize.HeightPu - pageSettings.Margins.TopPu - pageSettings.Margins.BottomPu;
                    y = pageSettings.Margins.BottomPu + (h - page.ChildObjects[0].Size.Y) / 2;
                }
                page.ChildObjects[0].LocalPosition = new Vector2(x, y);
                page.ContentTop = page.ChildObjects[0].LocalPosition.Y + page.ChildObjects[0].Size.Y;
                page.ContentBottom = page.ChildObjects[0].LocalPosition.Y;
                page.ContentLeft = page.ChildObjects[0].LocalPosition.X;
                page.ContentRight = page.ChildObjects[0].LocalPosition.X + page.ChildObjects[0].Size.X;
                page.ContentHeight = page.ChildObjects[0].Size.Y;

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
                                if (page.Map[r, c].content != null) page.Map[r, c].content.Clip = true;
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

                        // ── Text-spill propagation ───────────────────────────────────
                        if (leftCell.cell != null)
                        {
                            if (leftCell.content != null)
                            {
                                if (leftCell.content.RightTextSpillLength > 0d)
                                {
                                    if (cell.content != null)
                                    {
                                        // Spill collides with content — clip the source and force a gridline
                                        leftCell.content.Clip = true;
                                        addVertical(row, col);
                                    }
                                    else
                                    {
                                        cell.RightTextBucketSpill = leftCell.content.RightTextSpillLength - cell.cell.Size.X;
                                    }
                                }
                            }
                            else if (leftCell.RightTextBucketSpill > 0d)
                            {
                                if (cell.content != null)
                                {
                                    // Bucket spill collides with content — walk back to find origin and clip it, force a gridline
                                    for (int i = col - 1; i >= 0; i--)
                                    {
                                        var originCell = page.Map[row, i];
                                        if (originCell.content != null)
                                        {
                                            originCell.content.Clip = true;
                                            break;
                                        }
                                    }
                                    addVertical(row, col);
                                }
                                else
                                {
                                    cell.RightTextBucketSpill = leftCell.RightTextBucketSpill - cell.cell.Size.X;
                                }
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
                                        // Spill collides with content — clip and force a gridline at the boundary
                                        cell.content.Clip = true;
                                        cell.content.CreateClippingRect(cell.cell, prevCell.cell.LocalPosition.X + prevCell.cell.Size.X);
                                        addVertical(row, i);
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

                            // Border data is only stored at the merged cell's origin row — walk up to find it
                            int topOriginRow = row - 1;
                            while (topOriginRow > 0 && page.Map[topOriginRow - 1, col].cell == t.cell)
                                topOriginRow--;
                            var topBorder = page.Map[topOriginRow, col].border;

                            int cellOriginRow = row;
                            while (cellOriginRow > 0 && page.Map[cellOriginRow - 1, col].cell == cell.cell)
                                cellOriginRow--;
                            var cellBorder = page.Map[cellOriginRow, col].border;

                            bool borderTop = (topBorder != null && topBorder.BorderData.Bottom.BorderStyle != ExcelBorderStyle.None) ||
                                             (cellBorder != null && cellBorder.BorderData.Top.BorderStyle != ExcelBorderStyle.None);

                            if (differentTop && !borderTop)
                                AddHorizontalSegment(row - 1, col);
                        }

                        // Interior vertical — only between two columns (skip left page edge)
                        if (col > 0)
                        {
                            var l = page.Map[row, col - 1];
                            bool differentLeft = l.cell != cell.cell;
                            bool spillLeft = l.RightTextBucketSpill > 0 || cell.LeftTextBucketSpill > 0;

                            // Border data is only stored at the merged cell's origin column — walk left to find it
                            int leftOriginCol = col - 1;
                            while (leftOriginCol > 0 && page.Map[row, leftOriginCol - 1].cell == l.cell)
                                leftOriginCol--;
                            var leftBorder = page.Map[row, leftOriginCol].border;

                            int cellOriginCol = col;
                            while (cellOriginCol > 0 && page.Map[row, cellOriginCol - 1].cell == cell.cell)
                                cellOriginCol--;
                            var cellBorderV = page.Map[row, cellOriginCol].border;

                            bool borderLeft = (leftBorder != null && leftBorder.BorderData.Right.BorderStyle != ExcelBorderStyle.None) ||
                                              (cellBorderV != null && cellBorderV.BorderData.Left.BorderStyle != ExcelBorderStyle.None);

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

                // ── Border lines (4 outer edges, derived from the corner cells) ──────
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
                page.ChildObjects.RemoveAll(x => x is PdfCellLayout && ((PdfCellLayout)x).delete == true);
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
                    else if (child is PdfCellContentLayout content)
                    {
                        content.CreateClippingRect(page.ChildObjects);
                    }
                    else if (child is PdfMergedCellLayout mergedLayout)
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
                lh.Name = "LeftHeader";
                lh.LocalPosition = new Vector2(pageSettings.Margins.LeftPu, pageSettings.PageSize.HeightPu - pageSettings.Margins.HeaderPu);
                lh.AdjustPositionByTextLength('l', 'h');
                var ch = (PdfHeaderFooterLayout)pages.ChildObjects[i].AddChild(new PdfHeaderFooterLayout(centerH, ws, pageSettings, dictionaries, pageNumber, pages.ChildObjects.Count));
                ch.Name = "CenterHeader";
                ch.LocalPosition = new Vector2(pageSettings.PageSize.WidthPu / 2d, pageSettings.PageSize.HeightPu - pageSettings.Margins.HeaderPu);
                ch.AdjustPositionByTextLength('c', 'h');
                var rh = (PdfHeaderFooterLayout)pages.ChildObjects[i].AddChild(new PdfHeaderFooterLayout(rightH, ws, pageSettings, dictionaries, pageNumber, pages.ChildObjects.Count));
                rh.Name = "RightHeader";
                rh.LocalPosition = new Vector2(pageSettings.PageSize.WidthPu - pageSettings.Margins.RightPu, pageSettings.PageSize.HeightPu - pageSettings.Margins.HeaderPu);
                rh.AdjustPositionByTextLength('r', 'h');
                var lf = (PdfHeaderFooterLayout)pages.ChildObjects[i].AddChild(new PdfHeaderFooterLayout(leftF, ws, pageSettings, dictionaries, pageNumber, pages.ChildObjects.Count));
                lf.Name = "LeftFooter";
                lf.LocalPosition = new Vector2(pageSettings.Margins.LeftPu, pageSettings.Margins.FooterPu);
                var cf = (PdfHeaderFooterLayout)pages.ChildObjects[i].AddChild(new PdfHeaderFooterLayout(centerF, ws, pageSettings, dictionaries, pageNumber, pages.ChildObjects.Count));
                cf.Name = "CenterFooter";
                cf.LocalPosition = new Vector2(pageSettings.PageSize.WidthPu / 2d, pageSettings.Margins.FooterPu);
                cf.AdjustPositionByTextLength('c', 'f');
                var rf = (PdfHeaderFooterLayout)pages.ChildObjects[i].AddChild(new PdfHeaderFooterLayout(rightF, ws, pageSettings, dictionaries, pageNumber, pages.ChildObjects.Count));
                rf.Name = "RightFooter";
                rf.LocalPosition = new Vector2(pageSettings.PageSize.WidthPu - pageSettings.Margins.RightPu, pageSettings.Margins.FooterPu);
                rf.AdjustPositionByTextLength('r', 'f');
                pageNumber++;
            }
        }
    }
}
