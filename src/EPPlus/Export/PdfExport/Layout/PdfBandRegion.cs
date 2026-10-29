using EPPlus.Export.Pdf.Layout;
using OfficeOpenXml.Export.PdfExport.Data;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Export.PdfExport.Layout
{
    internal class PdfBandRegion
    {
        internal static void ProcessBandRegionCells(Page page, PdfRange range, int fromRow, int toRow, int fromCol, int toCol, Dictionary<int, double> colX, Dictionary<int, double> rowY)
        {
            double regLeft = colX[fromCol];
            double regRight = colX[toCol] + RangeColWidth(range, page.FromRow, toCol);
            double regTop = rowY[fromRow];
            double regBottom = rowY[toRow] - RangeRowHeight(range, toRow);
            double regW = regRight - regLeft;
            double regH = regTop - regBottom;
            var drawnMerges = new HashSet<string>();
            for (int r = fromRow; r <= toRow; r++)
            {
                for (int c = fromCol; c <= toCol; c++)
                {
                    var cell = RangeCell(range, r, c);
                    if (cell == null) continue;
                    if (cell.Merged)
                    {
                        if (!drawnMerges.Add(cell.MergedAddress.Address)) continue;
                        var addr = cell.MergedAddress;
                        var main = cell.Main ?? cell;
                        int vFromCol = Math.Max(addr._fromCol, fromCol);
                        int vToCol = Math.Min(addr._toCol, toCol);
                        int vFromRow = Math.Max(addr._fromRow, fromRow);
                        int vToRow = Math.Min(addr._toRow, toRow);
                        if (vFromCol > vToCol || vFromRow > vToRow) continue;
                        double x = colX[vFromCol];
                        double width = (colX[vToCol] + RangeColWidth(range, page.FromRow, vToCol)) - colX[vFromCol];
                        double y = rowY[vFromRow];
                        double height = (rowY[vFromRow] - rowY[vToRow]) + RangeRowHeight(range, vToRow);
                        if (width <= 0d || height <= 0d) continue;
                        page.PrintTitleCells.Add(new PrintTitleCellDraw
                        {
                            Cell = main,
                            X = x,
                            Y = y,
                            Width = width,
                            Height = height,
                            ClipX = x,
                            ClipY = y,
                            ClipWidth = width,
                            ClipHeight = height
                        });
                        for (int br = vFromRow; br <= vToRow; br++)
                        {
                            double bh = RangeRowHeight(range, br);
                            if (bh <= 0d) continue;
                            for (int bc = vFromCol; bc <= vToCol; bc++)
                            {
                                double bw = RangeColWidth(range, page.FromRow, bc);
                                if (bw <= 0d) continue;
                                var sub = RangeCell(range, br, bc);
                                if (sub == null || !HasBorder(sub.CellStyle)) continue;
                                page.PrintTitleBorders.Add(new PrintTitleCellDraw
                                {
                                    Cell = sub,
                                    X = colX[bc],
                                    Y = rowY[br],
                                    Width = bw,
                                    Height = bh
                                });
                            }
                        }
                    }
                    else
                    {
                        double w = RangeColWidth(range, page.FromRow, c);
                        double h = RangeRowHeight(range, r);
                        if (w <= 0d || h <= 0d) continue;
                        page.PrintTitleCells.Add(new PrintTitleCellDraw
                        {
                            Cell = cell,
                            X = colX[c],
                            Y = rowY[r],
                            Width = w,
                            Height = h,
                            ClipX = regLeft,
                            ClipY = regTop,
                            ClipWidth = regW,
                            ClipHeight = regH
                        });
                    }
                }
            }
        }

        internal static void EmitBandGrid(List<GridLine> target, PdfRange range, int fromRow, int toRow, int fromCol, int toCol, double[] colX, double[] rowY)
        {
            int nr = toRow - fromRow + 1;
            int nc = toCol - fromCol + 1;
            if (nr <= 0 || nc <= 0) return;
            var spill = BuildBandSpillMask(range, fromRow, toRow, fromCol, toCol, colX);
            double top = rowY[0], bottom = rowY[nr];
            double left = colX[0], right = colX[nc];
            EmitBandFrameH(target, range, top, fromRow, fromCol, nc, colX, CellHasTopBorder);
            EmitBandFrameH(target, range, bottom, toRow, fromCol, nc, colX, CellHasBottomBorder);
            EmitBandFrameV(target, range, left, fromCol, fromRow, nr, rowY, CellHasLeftBorder);
            EmitBandFrameV(target, range, right, toCol, fromRow, nr, rowY, CellHasRightBorder);
            for (int gi = 1; gi < nc; gi++)
            {
                int leftCol = fromCol + gi - 1, rightCol = leftCol + 1;
                double x = colX[gi];
                double? runStart = null; double runEnd = 0d;
                for (int ri = 0; ri < nr; ri++)
                {
                    int r = fromRow + ri;
                    bool block = SameMerge(range, r, leftCol, r, rightCol) || spill[ri, gi] ||
                                            CellHasRightBorder(RangeCell(range, r, leftCol)) ||
                                            CellHasLeftBorder(RangeCell(range, r, rightCol));
                    if (!block) { if (runStart == null) runStart = rowY[ri]; runEnd = rowY[ri + 1]; }
                    else if (runStart != null) { target.Add(new GridLine(x, runStart.Value, x, runEnd)); runStart = null; }
                }
                if (runStart != null) target.Add(new GridLine(x, runStart.Value, x, runEnd));
            }
            for (int gj = 1; gj < nr; gj++)
            {
                int topRow = fromRow + gj - 1, bottomRow = topRow + 1;
                double y = rowY[gj];
                double? runStart = null; double runEnd = 0d;
                for (int ci = 0; ci < nc; ci++)
                {
                    int c = fromCol + ci;
                    bool block = SameMerge(range, topRow, c, bottomRow, c) ||
                                            CellHasBottomBorder(RangeCell(range, topRow, c)) ||
                                            CellHasTopBorder(RangeCell(range, bottomRow, c));
                    if (!block) { if (runStart == null) runStart = colX[ci]; runEnd = colX[ci + 1]; }
                    else if (runStart != null) { target.Add(new GridLine(runStart.Value, y, runEnd, y)); runStart = null; }
                }
                if (runStart != null) target.Add(new GridLine(runStart.Value, y, runEnd, y));
            }
        }

        internal static bool[,] BuildBandSpillMask(PdfRange range, int fromRow, int toRow, int fromCol, int toCol, double[] colX)
        {
            int nr = toRow - fromRow + 1;
            int nc = toCol - fromCol + 1;
            var blocked = new bool[Math.Max(nr, 1), Math.Max(nc, 1)];
            if (nr <= 0 || nc <= 1) return blocked;
            int repRow = fromRow;
            for (int ri = 0; ri < nr; ri++)
            {
                int row = fromRow + ri;
                for (int ci = 0; ci < nc; ci++)
                {
                    var cell = RangeCell(range, row, fromCol + ci);
                    if (cell == null || cell.Merged) continue;
                    if (cell.ContentAligmnet != null && cell.ContentAligmnet.WrapText) continue;
                    if (string.IsNullOrEmpty(cell.Text)) continue;
                    double spill = cell.TotalTextLength - cell.ColumnWidth;
                    if (spill <= 0d) continue;
                    var hal = cell.ContentAligmnet?.HorizontalAlignment ?? (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.General;
                    if (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Left || hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.General)
                        BandMarkRight(range, row, ci, nc, fromCol, colX, spill, ri, blocked);
                    else if (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Right)
                        BandMarkLeft(range, row, ci, fromCol, colX, spill, ri, blocked);
                    else if (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Center)
                    {
                        double half = spill / 2d;
                        BandMarkRight(range, row, ci, nc, fromCol, colX, half, ri, blocked);
                        BandMarkLeft(range, row, ci, fromCol, colX, half, ri, blocked);
                    }
                }
                double lx = colX[0];
                for (int c = fromCol - 1; c >= range.Range._fromCol; c--)
                {
                    double w = RangeColWidth(range, repRow, c);
                    lx -= w;
                    var cell = RangeCell(range, row, c);
                    if (cell == null) continue;
                    if (cell.Merged) break;
                    if (string.IsNullOrEmpty(cell.Text)) continue;
                    var hal = cell.ContentAligmnet?.HorizontalAlignment ?? (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.General;
                    if (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Right) break;
                    double rightExtent = (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Center)
                        ? lx + w / 2d + cell.TotalTextLength / 2d
                        : lx + cell.TotalTextLength;
                    for (int g = 1; g <= nc - 1; g++)
                    {
                        var blk = RangeCell(range, row, fromCol + g - 1);
                        if (blk != null && !string.IsNullOrEmpty(blk.Text)) break;
                        if (colX[g] < rightExtent) blocked[ri, g] = true; else break;
                    }
                    break;
                }
                double rx = colX[nc];
                for (int c = toCol + 1; c <= range.Range._toCol; c++)
                {
                    double w = RangeColWidth(range, repRow, c);
                    var cell = RangeCell(range, row, c);
                    if (cell == null) { rx += w; continue; }
                    if (cell.Merged) break;
                    if (string.IsNullOrEmpty(cell.Text)) { rx += w; continue; }
                    var hal = cell.ContentAligmnet?.HorizontalAlignment ?? (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.General;
                    if (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Left || hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.General) break; // spills away
                    double leftExtent = (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Center)
                        ? rx + w / 2d - cell.TotalTextLength / 2d
                        : rx + w - cell.TotalTextLength;
                    for (int g = nc - 1; g >= 1; g--)
                    {
                        var blk = RangeCell(range, row, fromCol + g);
                        if (blk != null && !string.IsNullOrEmpty(blk.Text)) break;
                        if (colX[g] > leftExtent) blocked[ri, g] = true; else break;
                    }
                    break;
                }
            }
            return blocked;
        }

        internal static void BandMarkRight(PdfRange range, int row, int ci, int nc, int fromCol, double[] colX, double spill, int ri, bool[,] blocked)
        {
            for (int g = ci + 1; g <= nc - 1; g++)
            {
                var rightCell = RangeCell(range, row, fromCol + g);
                if (rightCell != null && !string.IsNullOrEmpty(rightCell.Text)) break;
                double distToGap = colX[g] - colX[ci + 1];
                if (spill > distToGap) blocked[ri, g] = true; else break;
            }
        }

        internal static void BandMarkLeft(PdfRange range, int row, int ci, int fromCol, double[] colX, double spill, int ri, bool[,] blocked)
        {
            for (int g = ci; g >= 1; g--)
            {
                var leftCell = RangeCell(range, row, fromCol + g - 1);
                if (leftCell != null && !string.IsNullOrEmpty(leftCell.Text)) break;
                double distToGap = colX[ci] - colX[g];
                if (spill > distToGap) blocked[ri, g] = true; else break;
            }
        }

        internal static bool SameMerge(PdfRange range, int r1, int c1, int r2, int c2)
        {
            var a = RangeCell(range, r1, c1);
            var b = RangeCell(range, r2, c2);
            if (a == null || b == null || !a.Merged || !b.Merged) return false;
            return a.MergedAddress.Address == b.MergedAddress.Address;
        }

        internal static double[] BandEdgesX(Dictionary<int, double> leftX, int fromCol, int toCol, PdfRange range, int repRow)
        {
            int n = toCol - fromCol + 1;
            var arr = new double[n + 1];
            for (int c = fromCol, k = 0; c <= toCol; c++, k++) arr[k] = leftX[c];
            arr[n] = leftX[toCol] + RangeColWidth(range, repRow, toCol);
            return arr;
        }

        internal static double[] BandEdgesY(Dictionary<int, double> topY, int fromRow, int toRow, PdfRange range)
        {
            int n = toRow - fromRow + 1;
            var arr = new double[n + 1];
            for (int r = fromRow, k = 0; r <= toRow; r++, k++) arr[k] = topY[r];
            arr[n] = topY[toRow] - RangeRowHeight(range, toRow);
            return arr;
        }

        internal static void EmitBandFrameH(List<GridLine> target, PdfRange range, double y, int row, int fromCol, int nc, double[] colX, Func<PdfCell, bool> hasBorder)
        {
            double? rs = null; double re = 0d;
            for (int ci = 0; ci < nc; ci++)
            {
                double segL = colX[ci], segR = colX[ci + 1];
                if (segR - segL <= 0d) { if (rs != null) { target.Add(new GridLine(rs.Value, y, re, y)); rs = null; } continue; }
                if (!hasBorder(RangeCell(range, row, fromCol + ci))) { if (rs == null) rs = segL; re = segR; }
                else if (rs != null) { target.Add(new GridLine(rs.Value, y, re, y)); rs = null; }
            }
            if (rs != null) target.Add(new GridLine(rs.Value, y, re, y));
        }

        internal static void EmitBandFrameV(List<GridLine> target, PdfRange range, double x, int col, int fromRow, int nr, double[] rowY, Func<PdfCell, bool> hasBorder)
        {
            double? rs = null; double re = 0d;
            for (int ri = 0; ri < nr; ri++)
            {
                double segT = rowY[ri], segB = rowY[ri + 1];
                if (segT - segB <= 0d) { if (rs != null) { target.Add(new GridLine(x, rs.Value, x, re)); rs = null; } continue; }
                if (!hasBorder(RangeCell(range, fromRow + ri, col))) { if (rs == null) rs = segT; re = segB; }
                else if (rs != null) { target.Add(new GridLine(x, rs.Value, x, re)); rs = null; }
            }
            if (rs != null) target.Add(new GridLine(x, rs.Value, x, re));
        }
    }
}
