using EPPlus.Export.Pdf.Settings;
using EPPlus.Export.Pdf.Settings.PdfPageSizes;
using OfficeOpenXml.Export.PdfExport.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.PdfExport.Layout
{
    internal class PdfPaginator
    {
        private struct PageSegment
        {
            public int From;
            public int To;
            public PageSegment(int from, int to) { From = from; To = to; }
        }

        internal static Pages GetNumberOfPages(PdfPageSettings pageSettings, PdfWorksheet pdfSheet, ref PdfRange range)
        {
            var xPages = (int)Math.Max(1, Math.Ceiling(range.TotalWidth / pageSettings.EffectiveContentWidth));
            var yPages = (int)Math.Max(1, Math.Ceiling(range.TotalHeight / pageSettings.EffectiveContentHeight));

            if (pageSettings.ShowHeadings)
            {
                int prev = 0;
                do
                {
                    prev = xPages;
                    range.AdditionalWidth = xPages * ((rowHeadingWith1CharWidth - pdfSheet.ZeroCharWidth) + (Math.Abs(pdfSheet.ToRow).ToString().Length * pdfSheet.ZeroCharWidth));
                    xPages = (int)Math.Max(1, Math.Ceiling((range.TotalWidth + range.AdditionalWidth) / pageSettings.EffectiveContentWidth));
                } while (prev != xPages);
                do
                {
                    prev = yPages;
                    range.AdditionalHeight = yPages * pdfSheet.Worksheet.DefaultRowHeight;
                    yPages = (int)Math.Max(1, Math.Ceiling((range.TotalHeight + range.AdditionalHeight) / pageSettings.EffectiveContentHeight));
                } while (prev != yPages);
            }
            if (pageSettings.Scaling?.ScalingMode != ScalingMode.FitToPages)
            {
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
            }
            ComputePrintTitleDimensions(pdfSheet, range, out range.PrintTitleHeight, out range.PrintTitleWidth);
            range.PrintTitleRowTo = pdfSheet.PrintTitleRowFrom >= 0 ? pdfSheet.PrintTitleRowTo : -1;
            range.PrintTitleColTo = pdfSheet.PrintTitleColFrom >= 0 ? pdfSheet.PrintTitleColTo : -1;
            Pages p = new Pages();
            p.Width = xPages;
            p.Height = yPages;
            p.Page = null;
            return p;
        }

        internal static Pages MapPage(PdfRange range, Pages pdfPages)
        {
            var pages = pdfPages;
            for (int i = 0; i < pdfPages.Page.Length; i++)
            {
                var page = pdfPages.Page[i];
                page.Map = new PdfCellCollection(page.FromRow, page.ToRow, page.FromColumn, page.ToColumn);
                page.RowHeights = new double[page.ToRow - page.FromRow + 1];
                for (int row = page.FromRow; row <= page.ToRow; row++)
                {
                    int rangeIdx = row - range.Range._fromRow;
                    page.RowHeights[row - page.FromRow] = range.RowHeights[rangeIdx].Height;
                    for (int col = page.FromColumn; col <= page.ToColumn; col++)
                    {
                        page.Map[row, col] = range.Map[row, col];
                    }
                }
                double usedWidth = page.HeadingWidth + page.PrintTitleWidth;
                for (int col = page.FromColumn; col <= page.ToColumn; col++)
                {
                    usedWidth += page.Map[page.FromRow, col]?.ColumnWidth ?? 0d;
                }
                page.UsedWidth = usedWidth;

                double usedHeight = page.HeadingHeight + page.PrintTitleHeight;
                for (int ri = 0; ri < page.RowHeights.Length; ri++)
                {
                    usedHeight += page.RowHeights[ri];
                }
                page.UsedHeight = usedHeight;

                pdfPages.Page[i] = page;
            }
            pdfPages = pages;
            return pdfPages;
        }

        private static List<PageSegment> GetColumnSegments(PdfPageSettings pageSettings, PdfRange range, ExcelWorksheet worksheet, double addedWidth, double titleWidth, int printTitleColTo)
        {
            var segments = new List<PageSegment>();
            int segStartIdx = 0;
            double width = 0d;
            for (int col = 0; col < range.ColWidths.Count; col++)
            {
                int actualCol = range.Range._fromCol + col;
                bool reserveTitle = titleWidth > 0d && printTitleColTo >= 0 && (range.Map.FromColumn + segStartIdx) > printTitleColTo;
                double effectiveAdded = addedWidth + (reserveTitle ? titleWidth : 0d);
                if (width + range.ColWidths[col] + effectiveAdded >= pageSettings.EffectiveContentWidth)
                {
                    if (col == segStartIdx)
                    {
                        segments.Add(new PageSegment(range.Map.FromColumn + col, range.Map.FromColumn + col));
                        segStartIdx = col + 1;
                        width = 0d;
                        continue;
                    }
                    segments.Add(new PageSegment(range.Map.FromColumn + segStartIdx, range.Map.FromColumn + col - 1));
                    segStartIdx = col;
                    width = 0d;
                    col--;
                    continue;
                }
                width += range.ColWidths[col];
                if (pageSettings.Scaling?.ScalingMode != ScalingMode.FitToPages && worksheet.Column(actualCol).PageBreak)
                {
                    segments.Add(new PageSegment(range.Map.FromColumn + segStartIdx, range.Map.FromColumn + col));
                    segStartIdx = col + 1;
                    width = 0d;
                }
            }
            if (segStartIdx < range.ColWidths.Count)
                segments.Add(new PageSegment(range.Map.FromColumn + segStartIdx, range.Map.FromColumn + range.ColWidths.Count - 1));
            return segments;
        }

        private static List<PageSegment> GetRowSegments(PdfPageSettings pageSettings, PdfRange range, ExcelWorksheet worksheet, double addedHeight, double titleHeight, int printTitleRowTo)
        {
            var segments = new List<PageSegment>();
            int segStartIdx = 0;
            double height = 0d;
            for (int row = 0; row < range.RowHeights.Count; row++)
            {
                int actualRow = range.Range._fromRow + row;
                bool reserveTitle = titleHeight > 0d && printTitleRowTo >= 0 && (range.Map.FromRow + segStartIdx) > printTitleRowTo;
                double effectiveAdded = addedHeight + (reserveTitle ? titleHeight : 0d);
                if (height + range.RowHeights[row].Height + effectiveAdded >= pageSettings.EffectiveContentHeight)
                {
                    segments.Add(new PageSegment(range.Map.FromRow + segStartIdx, range.Map.FromRow + row - 1));
                    segStartIdx = row;
                    height = 0d;
                    row--;
                    continue;
                }
                height += range.RowHeights[row].Height;
                if (pageSettings.Scaling?.ScalingMode != ScalingMode.FitToPages && worksheet.Row(actualRow).PageBreak)
                {
                    segments.Add(new PageSegment(range.Map.FromRow + segStartIdx, range.Map.FromRow + row));
                    segStartIdx = row + 1;
                    height = 0d;
                }
            }
            if (segStartIdx < range.RowHeights.Count)
                segments.Add(new PageSegment(range.Map.FromRow + segStartIdx, range.Map.FromRow + range.RowHeights.Count - 1));
            return segments;
        }

        internal static Pages AssignRangeToPages(PdfPageSettings pageSettings, PdfRange range, Pages pdfPages)
        {
            var pages = pdfPages;
            var worksheet = range.Range.Worksheet;
            var addedWidth = pages.Width > 0 ? range.AdditionalWidth / pages.Width : 0d;
            var addedHeight = pages.Height > 0 ? range.AdditionalHeight / pages.Height : 0d;
            var colSegments = GetColumnSegments(pageSettings, range, worksheet, addedWidth, range.PrintTitleWidth, range.PrintTitleColTo);
            var rowSegments = GetRowSegments(pageSettings, range, worksheet, addedHeight, range.PrintTitleHeight, range.PrintTitleRowTo);
            pages.Page = new Page[colSegments.Count * rowSegments.Count];
            int i = 0;

            if (pageSettings.PageOrders == PageOrders.DownThenOver)
            {
                for (int ci = 0; ci < colSegments.Count; ci++)
                    for (int ri = 0; ri < rowSegments.Count; ri++)
                        pages.Page[i++] = new Page
                        {
                            FromColumn = colSegments[ci].From,
                            ToColumn = colSegments[ci].To,
                            FromRow = rowSegments[ri].From,
                            ToRow = rowSegments[ri].To,
                            HeadingWidth = addedWidth,
                            HeadingHeight = addedHeight,
                            PrintTitleWidth = (range.PrintTitleColTo >= 0 && colSegments[ci].From > range.PrintTitleColTo) ? range.PrintTitleWidth : 0d,
                            PrintTitleHeight = (range.PrintTitleRowTo >= 0 && rowSegments[ri].From > range.PrintTitleRowTo) ? range.PrintTitleHeight : 0d,
                        };
            }
            else //if (pageSettings.PageOrders == PageOrders.OverThenDown)
            {
                for (int ri = 0; ri < rowSegments.Count; ri++)
                    for (int ci = 0; ci < colSegments.Count; ci++)
                        pages.Page[i++] = new Page
                        {
                            FromColumn = colSegments[ci].From,
                            ToColumn = colSegments[ci].To,
                            FromRow = rowSegments[ri].From,
                            ToRow = rowSegments[ri].To,
                            HeadingWidth = addedWidth,
                            HeadingHeight = addedHeight,
                            PrintTitleWidth = (range.PrintTitleColTo >= 0 && colSegments[ci].From > range.PrintTitleColTo) ? range.PrintTitleWidth : 0d,
                            PrintTitleHeight = (range.PrintTitleRowTo >= 0 && rowSegments[ri].From > range.PrintTitleRowTo) ? range.PrintTitleHeight : 0d,
                        };
            }
            pdfPages = pages;
            return pdfPages;
        }

        internal static void ResolveFitScale(PdfPageSettings pageSettings, PdfWorksheet pdfSheet, PdfRange range)
        {
            var scaling = pageSettings.Scaling;
            if (scaling == null || scaling.ScalingMode != ScalingMode.FitToPages) return;

            int N = scaling.PagesWide;   // 0 = automatic
            int M = scaling.PagesTall;
            var worksheet = range.Range.Worksheet;

            // Same per-page reservations the real pagination uses (both are per-page constants — no circularity).
            double addedWidth = pageSettings.ShowHeadings
                ? (rowHeadingWith1CharWidth - pdfSheet.ZeroCharWidth) + (System.Math.Abs(pdfSheet.ToRow).ToString().Length * pdfSheet.ZeroCharWidth)
                : 0d;
            double addedHeight = pageSettings.ShowHeadings ? pdfSheet.Worksheet.DefaultRowHeight : 0d;

            ComputePrintTitleDimensions(pdfSheet, range, out double titleHeight, out double titleWidth);
            int titleColTo = pdfSheet.PrintTitleColFrom >= 0 ? pdfSheet.PrintTitleColTo : -1;
            int titleRowTo = pdfSheet.PrintTitleRowFrom >= 0 ? pdfSheet.PrintTitleRowTo : -1;

            bool Fits(double s)
            {
                pageSettings.SetResolvedFitScale(s);
                bool okW = N <= 0 || GetColumnSegments(pageSettings, range, worksheet, addedWidth, titleWidth, titleColTo).Count <= N;
                bool okH = M <= 0 || GetRowSegments(pageSettings, range, worksheet, addedHeight, titleHeight, titleRowTo).Count <= M;
                return okW && okH;
            }

            if (Fits(1.0d)) { pageSettings.SetResolvedFitScale(1.0d); return; }   // fits at 100% → don't shrink

            // Largest scale in [0.1, 1.0] whose real pagination fits N×M (monotonic: fewer pages as scale shrinks).
            double lo = 0.1d, hi = 1.0d;
            for (int i = 0; i < 50; i++)
            {
                double mid = (lo + hi) * 0.5d;
                if (Fits(mid)) lo = mid; else hi = mid;
            }
            pageSettings.SetResolvedFitScale(lo);
        }

        private static void ComputePrintTitleDimensions(PdfWorksheet pdfSheet, PdfRange range, out double titleHeight, out double titleWidth)
        {
            titleHeight = 0d;
            titleWidth = 0d;
            if (pdfSheet.PrintTitleRowFrom >= 0)
            {
                for (int r = pdfSheet.PrintTitleRowFrom; r <= pdfSheet.PrintTitleRowTo; r++)
                {
                    int idx = r - range.Range._fromRow;
                    if (idx >= 0 && idx < range.RowHeights.Count)
                        titleHeight += range.RowHeights[idx].Height;
                }
            }
            if (pdfSheet.PrintTitleColFrom >= 0)
            {
                for (int c = pdfSheet.PrintTitleColFrom; c <= pdfSheet.PrintTitleColTo; c++)
                {
                    int idx = c - range.Range._fromCol;
                    if (idx >= 0 && idx < range.ColWidths.Count)
                        titleWidth += range.ColWidths[idx];
                }
            }
        }
    }
}
