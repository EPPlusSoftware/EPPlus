using EPPlus.Export.Pdf.Enums;
using EPPlus.Export.Pdf.Layout;
using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings;
using OfficeOpenXml.Export.PdfExport.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.PdfExport.Layout
{
    internal class PdfSpillLayout
    {
        internal static Pages PrecomputeSpillCells(PdfPageSettings pageSettings, PdfRange range, Pages pdfPages)
        {
            for (int i = 0; i < pdfPages.Page.Length; i++)
            {
                var page = pdfPages.Page[i];
                page.SpillCells = new List<SpillCellDraw>();
                double originX = GetOriginX(pageSettings, page) + page.HeadingWidth + page.PrintTitleWidth;
                double originY = GetOriginY(pageSettings, page) - page.HeadingHeight - page.PrintTitleHeight;
                AddIncomingSpill(page, range, page.FromRow, page.ToRow, page.FromColumn, page.ToColumn, originX, originY, isPrintTitle: false);
                pdfPages.Page[i] = page;
            }
            return pdfPages;
        }

        private static void AddIncomingSpill(Page page, PdfRange range, int fromRow, int toRow, int windowFromCol, int windowToCol, double windowOriginX, double windowOriginY, bool isPrintTitle)
        {
            if (page.SpillCells == null) return;
            double windowRightX = windowOriginX;
            for (int c = windowFromCol; c <= windowToCol; c++) windowRightX += RangeColWidth(range, page.FromRow, c);
            double rowTop = windowOriginY;
            for (int r = fromRow; r <= toRow; r++)
            {
                double rowH = RangeRowHeight(range, r);
                double y = rowTop;
                rowTop -= rowH;
                if (rowH <= 0d) continue;
                double lx = windowOriginX;
                for (int c = windowFromCol - 1; c >= range.Range._fromCol; c--)
                {
                    double w = RangeColWidth(range, page.FromRow, c);
                    lx -= w;
                    var cell = RangeCell(range, r, c);
                    if (cell == null) continue;
                    if (cell.Merged) break;
                    if (string.IsNullOrEmpty(cell.Text)) continue;
                    var hal = cell.ContentAligmnet?.HorizontalAlignment ?? (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.General;
                    double rightExtent =
                        (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Center) ? lx + w / 2d + cell.TotalTextLength / 2d :
                        (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Right) ? lx + w : lx + cell.TotalTextLength;
                    if (rightExtent > windowOriginX)
                    {
                        double clipRight = FirstBlockedX(page, range, r, windowFromCol, windowToCol, windowOriginX, windowRightX, fromLeft: true);
                        page.SpillCells.Add(new SpillCellDraw
                        {
                            Cell = cell,
                            X = lx,
                            Y = y,
                            Width = w,
                            Height = rowH,
                            ClipX = windowOriginX,
                            ClipY = y,
                            ClipWidth = clipRight - windowOriginX,
                            ClipHeight = rowH,
                            IsPrintTitle = isPrintTitle
                        });
                    }
                    break;
                }
                double rx = windowRightX;
                for (int c = windowToCol + 1; c <= range.Range._toCol; c++)
                {
                    double w = RangeColWidth(range, page.FromRow, c);
                    var cell = RangeCell(range, r, c);
                    if (cell == null) { rx += w; continue; }
                    if (cell.Merged) break;
                    if (string.IsNullOrEmpty(cell.Text)) { rx += w; continue; }
                    var hal = cell.ContentAligmnet?.HorizontalAlignment ?? (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.General;
                    double leftExtent =
                        (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Center) ? rx + w / 2d - cell.TotalTextLength / 2d :
                        (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Right) ? rx + w - cell.TotalTextLength : rx;
                    if (leftExtent < windowRightX)
                    {
                        double clipLeft = FirstBlockedX(page, range, r, windowFromCol, windowToCol, windowOriginX, windowRightX, fromLeft: false);
                        page.SpillCells.Add(new SpillCellDraw
                        {
                            Cell = cell,
                            X = rx,
                            Y = y,
                            Width = w,
                            Height = rowH,
                            ClipX = clipLeft,
                            ClipY = y,
                            ClipWidth = windowRightX - clipLeft,
                            ClipHeight = rowH,
                            IsPrintTitle = isPrintTitle
                        });
                    }
                    break;
                }
            }
        }

        private static double FirstBlockedX(Page page, PdfRange range, int row, int windowFromCol, int windowToCol, double windowOriginX, double windowRightX, bool fromLeft)
        {
            if (fromLeft)
            {
                double x = windowOriginX;
                for (int c = windowFromCol; c <= windowToCol; c++)
                {
                    var cell = RangeCell(range, row, c);
                    if (cell != null && (cell.Merged || !string.IsNullOrEmpty(cell.Text))) return x;
                    x += RangeColWidth(range, page.FromRow, c);
                }
                return windowRightX;
            }
            double rx = windowRightX;
            for (int c = windowToCol; c >= windowFromCol; c--)
            {
                double w = RangeColWidth(range, page.FromRow, c);
                rx -= w;
                var cell = RangeCell(range, row, c);
                if (cell != null && (cell.Merged || !string.IsNullOrEmpty(cell.Text))) return rx + w;
            }
            return windowOriginX;
        }

        internal static void AddSpillCells(PdfPageSettings pageSettings, PdfDictionaries dictionaries, Page page, PdfPageLayout pageLayout)
        {
            if (page.SpillCells == null) return;
            foreach (var s in page.SpillCells)
            {
                var map = s.Cell;
                if (map.TextLines == null || map.TextLines.Count == 0) continue;
                if (s.ClipWidth <= 0d || s.ClipHeight <= 0d) continue;

                var text = new PdfCellContentLayout(pageSettings, dictionaries, map, s.X, s.Y, s.Width, s.Height);
                text.Name = "Spill_" + map.Name;
                text.IsPrintTitle = s.IsPrintTitle;
                text.GidsAndCharMap(dictionaries);
                text.SetupClipping(s.ClipX, s.ClipY, s.ClipWidth, s.ClipHeight);
                pageLayout.AddChild(text);
            }
        }

        internal static (double clipX, double clipW, bool spillLeft, bool spillRight) GetOverflowClip(PdfCell map, Page page, int row, int col, double x, double effectiveWidth)
        {
            double clipX = x, clipW = effectiveWidth;
            var a = map.ContentAligmnet;
            if (a == null)
                return (clipX, clipW, false, false);
            var L = (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Left;
            var G = (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.General;
            var R = (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Right;
            var C = (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Center;
            var F = (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Fill;
            var h = a.HorizontalAlignment;
            if (h == F || a.WrapText || map.TotalTextLength <= map.ColumnWidth)
                return (clipX, clipW, false, false);
            double ScanRight()
            {
                double w = 0d;
                for (int c = col + 1; c <= page.ToColumn; c++)
                {
                    var n = page.Map[row, c];
                    if (!string.IsNullOrEmpty(n.Text)) break;
                    w += n.ColumnWidth;
                }
                return w;
            }
            double ScanLeft()
            {
                double w = 0d;
                for (int c = col - 1; c >= page.FromColumn; c--)
                {
                    var n = page.Map[row, c];
                    if (!string.IsNullOrEmpty(n.Text)) break;
                    w += n.ColumnWidth;
                }
                return w;
            }
            if (h == L || h == G)
            {
                clipW = effectiveWidth + ScanRight();
                return (clipX, clipW, false, true);
            }
            if (h == R)
            {
                double extra = ScanLeft();
                return (x - extra, effectiveWidth + extra, true, false);
            }
            if (h == C)
            {
                double left = ScanLeft(), right = ScanRight();
                return (x - left, effectiveWidth + left + right, true, true);
            }
            return (clipX, clipW, false, false);
        }
    }
}
