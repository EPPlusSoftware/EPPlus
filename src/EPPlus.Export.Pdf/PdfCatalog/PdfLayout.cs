using EPPlus.Export.Pdf.PdfLayout;
using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Graphics;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.Pdf.PdfCatalog
{

    internal struct MergedCellDrawInfo
    {
        public double X;
        public double Y;
        public double Width;
        public double Height;
    }

    internal struct Page
    {
        public int FromRow;
        public int FromColumn;
        public int ToRow;
        public int ToColumn;

        public bool HasPrintTitle;

        public PdfCellCollection Map;

        public Dictionary<string, MergedCellDrawInfo> MergedCells;
    }

    internal struct Pages
    {
        public Page[] Page;
        public int Width;
        public int Height;
        public int Count
        {
            get { return Width * Height; }
        }
    }

    internal class PdfLayout
    {
        private const double rowHeadingWith1CharWidth = 23.25d;

        public static Transform GetLayout(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfWorksheet[] pdfSheets)
        {
            var PagesCollection = GetPages(pageSettings, pdfSheets);
            var Catalog = GetCatalog(pageSettings, dictionaries, PagesCollection);
            return Catalog;
        }

        internal static Transform GetCatalog(PdfPageSettings pageSettings, PdfDictionaries dictionaries, List<Pages> pdfPages)
        {
            Transform Catalog = new Transform(0d, 0d, 0d, 0d);
            int totalPages = GetTotalPages(pdfPages);
            for (int i = 0; i < pdfPages.Count; i++)
            {

                var pages = pdfPages[i].Page;
                int pageNumber = pageSettings.FirstPageNumber;

                for (int j = 0; j < pages.Length; j++)
                {
                    PdfPageLayout pageLayout = new PdfPageLayout(0d, 0d, 0d, 0d);
                    var drawnMergedCells = new HashSet<string>();
                    //var drawnMergedCellsText = new HashSet<string>();
                    //PdfContentLayout contentLayout = new PdfContentLayout(0d, 0d, pageSettings.ContentBounds);
                    //pageLayout.AddChild(contentLayout);
                    double y = pageSettings.ContentBounds.Top;
                    double x = pageSettings.ContentBounds.Left;
                    //create cells & headings if exsists
                    for (int row = pages[j].FromRow; row <= pages[j].ToRow; row++)
                    {
                        for (int col = pages[j].FromColumn; col <= pages[j].ToColumn; col++)
                        {
                            var map = pages[j].Map[row, col];
                            MergedCellDrawInfo info = new MergedCellDrawInfo();
                            //  Fill
                            if (map.Merged)
                            {
                                string key = map.MergedAddress.Address;
                                if (!drawnMergedCells.Contains(key) &&
                                    pages[j].MergedCells.TryGetValue(key, out info))
                                {
                                    var cellStyle = map.Main?.CellStyle ?? map.CellStyle;
                                    var fill = new PdfCellLayout(dictionaries, cellStyle,
                                        info.X, info.Y, info.Width, info.Height);
                                    fill.Name = map.Name;
                                    fill.UpdateShadingPositionMatrix(pageSettings);
                                    pageLayout.AddChild(fill);
                                    if (map.TextLines != null && map.TextLines.Count > 0)
                                    {
                                        var text = new PdfCellContentLayout(pageSettings, dictionaries, map, info, x, y, map.ColumnWidth, 15);
                                        text.Name = map.Name;
                                        text.GidsAndCharMap(dictionaries);
                                        pageLayout.AddChild(text);
                                    }
                                    else if (map.Main != null && map.Main.TextLines != null && map.Main.TextLines.Count > 0)
                                    {
                                        var text = new PdfCellContentLayout(pageSettings, dictionaries, map.Main, info, x, y, map.Main.ColumnWidth, 15);
                                        text.Name = map.Main.Name;
                                        text.GidsAndCharMap(dictionaries);
                                        pageLayout.AddChild(text);
                                    }
                                    drawnMergedCells.Add(key);
                                }
                            }
                            else
                            {
                                var fill = new PdfCellLayout(dictionaries, map.CellStyle, x, y, map.ColumnWidth, 15);
                                fill.UpdateShadingPositionMatrix(pageSettings);
                                fill.Name = map.Name;
                                pageLayout.AddChild(fill);
                            }
                            //Text
                            if (map.TextLines != null && map.TextLines.Count > 0)
                            {
                                var text = new PdfCellContentLayout(pageSettings, dictionaries, map, info, x, y, map.ColumnWidth, 15);
                                text.Name = map.Name;
                                text.GidsAndCharMap(dictionaries);
                                pageLayout.AddChild(text);
                            }
                            //Border
                            if (HasBorder(map.CellStyle))
                            {
                                var border = new PdfCellBorderLayout(map.CellStyle, map.Merged, GetCorners(map.MergedAddress, row, col), info, x, y, map.ColumnWidth, 15);
                                border.Name = map.Name;
                                pageLayout.AddChild(border);
                            }


                            x += map.ColumnWidth;
                        }
                        y -= 15;
                        x = pageSettings.ContentBounds.Left;
                    }


                    //Add HeaderFooter
                    //  Uppdate page number texts and shape them

                    //Gridlines (calculate text spill here)


                    //Print titles
                    pageNumber++;
                    Catalog.AddChild(pageLayout);
                }
            }
            return Catalog;
        }


        static MergedCellCorners GetCorners(ExcelAddressBase addr, int row, int col)
        {
            if (addr == null) return MergedCellCorners.All;
            MergedCellCorners result = MergedCellCorners.None;

            if (row == addr.Start.Row && col == addr.Start.Column)
                result |= MergedCellCorners.TopLeft;

            if (row == addr.Start.Row && col == addr.End.Column)
                result |= MergedCellCorners.TopRight;

            if (row == addr.End.Row && col == addr.Start.Column)
                result |= MergedCellCorners.BottomLeft;

            if (row == addr.End.Row && col == addr.End.Column)
                result |= MergedCellCorners.BottomRight;

            return result;
        }

        private static bool HasBorder(PdfCellStyle cellStyle)
        {
            if(cellStyle == null) return false;
            bool hasBorders =
                cellStyle.xfTop.Style != ExcelBorderStyle.None ||
                cellStyle.xfBottom.Style != ExcelBorderStyle.None ||
                cellStyle.xfLeft.Style != ExcelBorderStyle.None ||
                cellStyle.xfRight.Style != ExcelBorderStyle.None ||
                (cellStyle.dxfTop?.HasValue ?? false) ||
                (cellStyle.dxfBottom?.HasValue ?? false) ||
                (cellStyle.dxfLeft?.HasValue ?? false) ||
                (cellStyle.dxfRight?.HasValue ?? false) ||
                (cellStyle.Diagonal != null && cellStyle.Diagonal.Style != ExcelBorderStyle.None);
            return hasBorders;
        }

        // TODO Count total pages when creating them isntead of looping them here.
        private static int GetTotalPages(List<Pages> pdfPages)
        {
            int totalPages = 0;
            for (int i = 0; i < pdfPages.Count; i++)
            {
                totalPages += pdfPages[i].Page.Length;
            }
            return totalPages;
        }

        internal static List<Pages> GetPages(PdfPageSettings pageSettings, PdfWorksheet[] pdfSheets)
        {
            List<Pages> PagesCollection = new List<Pages>();
            foreach (var pdfSheet in pdfSheets)
            {
                foreach (var range in pdfSheet.Ranges)
                {
                    var pages = GetNumberOfPages(pageSettings, pdfSheet, range);
                    pages = AssignRangeToPages(pageSettings, range, pages);
                    pages = MapPage(range, pages);
                    pages = PrecomputeMergedCells(pageSettings, range, pages);
                    PagesCollection.Add(pages);
                }
                if (pdfSheet.CommentsAndNotes.Range != null)
                {
                    var pages = GetNumberOfPages(pageSettings, pdfSheet, pdfSheet.CommentsAndNotes);
                    pages = AssignRangeToPages(pageSettings, pdfSheet.CommentsAndNotes, pages);
                    pages = MapPage(pdfSheet.CommentsAndNotes, pages);
                    PagesCollection.Add(pages);
                }
            }
            return PagesCollection;
        }

        internal static Pages PrecomputeMergedCells(PdfPageSettings pageSettings, PdfRange range, Pages pdfPages)
        {
            for (int i = 0; i < pdfPages.Page.Length; i++)
                pdfPages.Page[i] = PrecomputePageMergedCells(pageSettings, range, pdfPages.Page[i]);
            return pdfPages;
        }

        private static Page PrecomputePageMergedCells(PdfPageSettings pageSettings, PdfRange range, Page page)
        {
            page.MergedCells = new Dictionary<string, MergedCellDrawInfo>();
            // Build a quick lookup: absolute x for each column index on this page.
            // We read ColumnWidth from the first data row; widths are per-column, not per-cell.
            var colX = BuildColumnXPositions(pageSettings, page);
            for (int row = page.FromRow; row <= page.ToRow; row++)
            {
                for (int col = page.FromColumn; col <= page.ToColumn; col++)
                {
                    var cell = page.Map[row, col];
                    if (cell == null || !cell.Merged) continue;
                    string key = cell.MergedAddress.Address;
                    if (page.MergedCells.ContainsKey(key)) continue;
                    var addr = cell.MergedAddress;
                    var mainCell = cell.Main ?? cell; // Main == null means this cell IS the top-left
                    // --- X ---
                    // Start from the current column and walk left to the merge origin.
                    // Columns within the current page come from colX; columns that lie on
                    // a preceding column-page come from range.ColWidths.
                    double drawX = colX[col - page.FromColumn];
                    for (int c = addr._fromCol; c < col; c++)
                    {
                        int rangeIdx = c - range.Range._fromCol;
                        if (rangeIdx >= 0 && rangeIdx < range.ColWidths.Count)
                            drawX -= range.ColWidths[rangeIdx];
                    }
                    // --- Y ---
                    // y decreases as row increases (PDF coordinate origin at bottom-left).
                    // Build the y for the current row, then add back the heights of rows
                    // above it that belong to the merge (whether on this page or a prior one).
                    double drawY = pageSettings.ContentBounds.Top - (row - page.FromRow) * 15d; // TODO: real row heights
                    for (int r = addr._fromRow; r < row; r++)
                    {
                        int rangeIdx = r - range.Range._fromRow;
                        if (rangeIdx >= 0 && rangeIdx < range.RowHeights.Count)
                            drawY += range.RowHeights[rangeIdx];
                    }
                    page.MergedCells[key] = new MergedCellDrawInfo
                    {
                        X = drawX,
                        Y = drawY,
                        Width = mainCell.Width,
                        Height = mainCell.Height
                    };
                }
            }
            return page;
        }

        private static double[] BuildColumnXPositions(PdfPageSettings pageSettings, Page page)
        {
            int colCount = page.ToColumn - page.FromColumn + 1;
            var colX = new double[colCount];
            double x = pageSettings.ContentBounds.Left;
            for (int col = page.FromColumn; col <= page.ToColumn; col++)
            {
                colX[col - page.FromColumn] = x;
                var cell = page.Map[page.FromRow, col];
                x += cell?.ColumnWidth ?? 0d;
            }
            return colX;
        }

        internal static Pages GetNumberOfPages(PdfPageSettings pageSettings, PdfWorksheet pdfSheet,  PdfRange range)
        {
            //calculte pages needed for this range, add in col headings for width, row headings for height. THis is where we also add print headings later on. Autofit on row here too later on.
            var xPages = (int)Math.Max(1, Math.Ceiling(range.TotalWidth / pageSettings.ContentBounds.Width));
            var yPages = (int)Math.Max(1, Math.Ceiling(range.TotalHeight / pageSettings.ContentBounds.Height));

            if (pageSettings.ShowHeadings)
            {
                int prev = 0;
                do
                {
                    prev = xPages;
                    range.AdditionalWidth = xPages * ((rowHeadingWith1CharWidth - pdfSheet.ZeroCharWidth) + (Math.Abs(pdfSheet.ToRow).ToString().Length * pdfSheet.ZeroCharWidth));
                    xPages = (int)Math.Max(1, Math.Ceiling((range.TotalWidth + range.AdditionalWidth) / pageSettings.ContentBounds.Width));
                } while (prev != xPages);
                do
                {
                    prev = yPages;
                    range.AdditionalHeight = yPages * pdfSheet.Worksheet.DefaultRowHeight;
                    yPages = (int)Math.Max(1, Math.Ceiling((range.TotalHeight + range.AdditionalHeight) / pageSettings.ContentBounds.Height));
                } while (prev != yPages);
            }
            for (int i = range.Range._fromCol; i <= range.Range._toCol; i++)
            {
                if (pdfSheet.Worksheet.Column(i).PageBreak)
                    xPages++;
            }
            for (int i = range.Range._fromRow; i <= range.Range._toRow; i++)
            {
                if (pdfSheet.Worksheet.Row(i).PageBreak)
                    yPages++;
            }
            //if (HasPrintTitles Row)
            //if (HasPrintTitles Column)

            Pages p;
            p.Width = xPages;
            p.Height = yPages;
            p.Page = null;
            return p;
        }

        internal static Pages AssignRangeToPages(PdfPageSettings pageSettings, PdfRange range, Pages pdfPages)
        {
            var pages = pdfPages;
            var worksheet = range.Range.Worksheet;
            var addedWidth = pages.Width > 0 ? range.AdditionalWidth / pages.Width : 0d;
            var addedHeight = pages.Height > 0 ? range.AdditionalHeight / pages.Height : 0d;

            var colSegments = GetColumnSegments(pageSettings, range, worksheet, addedWidth);
            var rowSegments = GetRowSegments(pageSettings, range, worksheet, addedHeight);

            pages.Page = new Page[colSegments.Count * rowSegments.Count];
            int i = 0;

            if (pageSettings.PageOrders == PageOrders.DownThenOver)
            {
                foreach (var colSeg in colSegments)
                    foreach (var rowSeg in rowSegments)
                        pages.Page[i++] = new Page { FromColumn = colSeg.From, ToColumn = colSeg.To, FromRow = rowSeg.From, ToRow = rowSeg.To };
            }
            else //if (pageSettings.PageOrders == PageOrders.OverThenDown)
            {
                foreach (var rowSeg in rowSegments)
                    foreach (var colSeg in colSegments)
                        pages.Page[i++] = new Page { FromColumn = colSeg.From, ToColumn = colSeg.To, FromRow = rowSeg.From, ToRow = rowSeg.To };
            }

            pdfPages = pages;
            return pdfPages;
        }

        private struct PageSegment
        {
            public int From;
            public int To;
            public PageSegment(int from, int to) { From = from; To = to; }
        }

        private static List<PageSegment> GetColumnSegments(PdfPageSettings pageSettings, PdfRange range, ExcelWorksheet worksheet, double addedWidth)
        {
            var segments = new List<PageSegment>();
            int segStartIdx = 0;
            double width = 0d;

            for (int col = 0; col < range.ColWidths.Count; col++)
            {
                int actualCol = range.Range._fromCol + col;

                // Content-bounds overflow: col doesn't fit, end segment before it and reprocess.
                if (width + range.ColWidths[col] + addedWidth >= pageSettings.ContentBounds.Width)
                {
                    segments.Add(new PageSegment(range.Map.FromColumn + segStartIdx, range.Map.FromColumn + col - 1));
                    segStartIdx = col;
                    width = 0d;
                    col--; // reprocess this col as the first col of the next segment
                    continue;
                }

                width += range.ColWidths[col];

                // Explicit page break: col is included on this page, next segment starts after it.
                if (worksheet.Column(actualCol).PageBreak)
                {
                    segments.Add(new PageSegment(range.Map.FromColumn + segStartIdx, range.Map.FromColumn + col));
                    segStartIdx = col + 1;
                    width = 0d;
                }
            }

            // Remaining cols form the last segment.
            if (segStartIdx < range.ColWidths.Count)
                segments.Add(new PageSegment(range.Map.FromColumn + segStartIdx, range.Map.FromColumn + range.ColWidths.Count - 1));

            return segments;
        }

        private static List<PageSegment> GetRowSegments(PdfPageSettings pageSettings, PdfRange range, ExcelWorksheet worksheet, double addedHeight)
        {
            var segments = new List<PageSegment>();
            int segStartIdx = 0;
            double height = 0d;

            for (int row = 0; row < range.RowHeights.Count; row++)
            {
                int actualRow = range.Range._fromRow + row;

                // Content-bounds overflow: row doesn't fit, end segment before it and reprocess.
                if (height + range.RowHeights[row] + addedHeight >= pageSettings.ContentBounds.Height)
                {
                    segments.Add(new PageSegment(range.Map.FromRow + segStartIdx, range.Map.FromRow + row - 1));
                    segStartIdx = row;
                    height = 0d;
                    row--; // reprocess this row as the first row of the next segment
                    continue;
                }

                height += range.RowHeights[row];

                // Explicit page break: row is included on this page, next segment starts after it.
                if (worksheet.Row(actualRow).PageBreak)
                {
                    segments.Add(new PageSegment(range.Map.FromRow + segStartIdx, range.Map.FromRow + row));
                    segStartIdx = row + 1;
                    height = 0d;
                }
            }

            // Remaining rows form the last segment.
            if (segStartIdx < range.RowHeights.Count)
                segments.Add(new PageSegment(range.Map.FromRow + segStartIdx, range.Map.FromRow + range.RowHeights.Count - 1));

            return segments;
        }

        internal static Pages MapPage(PdfRange range, Pages pdfPages)
        {
            var pages = pdfPages;
            for (int i = 0; i < pdfPages.Page.Length; i++)
            {
                var page = pdfPages.Page[i];
                page.Map = new PdfCellCollection(page.FromRow, page.ToRow, page.FromColumn, page.ToColumn);
                for (int row = page.FromRow; row <= page.ToRow; row++)
                {
                    for (int col = page.FromColumn; col <= page.ToColumn; col++)
                    {
                        page.Map[row, col] = range.Map[row, col];
                    }
                }
                pdfPages.Page[i] = page;
            }
            pdfPages = pages;
            return pdfPages;
        }

    }
}
