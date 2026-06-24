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
using EPPlus.Export.Pdf.Settings;
using OfficeOpenXml.Export.PdfExport.Data;
using EPPlus.Export.Pdf.Layout;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Export.PdfExport.Layout
{
    internal class PdfGridlinesLayout
    {
        //------------------------------------------------------------------------------------------
        // Gridline generation
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Populates <paramref name="pageLayout"/>.GridLines and .BorderLines for one page.
        ///
        /// Algorithm:
        ///   1. Precompute column X positions and row Y positions from the page map / row-height array.
        ///   2. Walk every cell and mark which vertical column-gaps are suppressed by text spill.
        ///      Spill direction depends on horizontal alignment; spill stops when a destination cell
        ///      has content.  Multi-column spill is handled by walking until either the text runs
        ///      out or a non-empty cell is found.
        ///   3. Emit vertical gridline segments: for each column gap, accumulate contiguous rows
        ///      that are not blocked by spill, a cell border, or a merged cell spanning the gap.
        ///   4. Emit horizontal gridline segments: for each row gap, accumulate contiguous columns
        ///      that are not blocked by a cell border or a merged cell spanning the gap.
        ///   5. Add the four outer-frame lines to BorderLines.
        /// </summary>
        public static void AddGridLines(PdfPageSettings pageSettings, Page page, PdfPageLayout pageLayout, bool borderOnly = false)
        {
            int rowCount = page.ToRow - page.FromRow + 1;
            int colCount = page.ToColumn - page.FromColumn + 1;

            if (rowCount <= 0 || colCount <= 0) return;

            // --- 1. Position arrays ---------------------------------------------------------------

            // colX[ci] = X of left edge of column ci (0-based within page).
            // colX[colCount] = X of right edge of last column.
            var colX = new double[colCount + 1];
            colX[0] = pageSettings.ContentBounds.Left + page.HeadingWidth + page.PrintTitleWidth;
            for (int ci = 0; ci < colCount; ci++)
            {
                var cell = page.Map[page.FromRow, page.FromColumn + ci];
                colX[ci + 1] = colX[ci] + (cell?.ColumnWidth ?? 0d);
            }

            // rowY[ri] = Y of top edge of row ri (0-based within page).
            // rowY[rowCount] = Y of bottom edge of last row.
            // Y decreases downward (PDF coordinate system used throughout GetCatalog).
            var rowY = new double[rowCount + 1];
            rowY[0] = pageSettings.ContentBounds.Top - page.HeadingHeight - page.PrintTitleHeight;

           for (int ri = 0; ri < rowCount; ri++)
            {
                rowY[ri + 1] = rowY[ri] - page.RowHeights[ri];
            }

            // --- 5. Outer frame ------------------------------------------------------------------
            // Always computed so BorderLines is available for margin clipping regardless of
            // whether ShowGridLines is on. When borderOnly is true we stop here.

            double frameLeft = pageSettings.ContentBounds.Left; //colX[0];
            double frameRight = colX[colCount];
            double frameTop = pageSettings.ContentBounds.Top; //rowY[0];
            double frameBottom = rowY[rowCount];

            pageLayout.BorderLines.Add(new GridLine(frameLeft, frameTop, frameRight, frameTop));
            pageLayout.BorderLines.Add(new GridLine(frameLeft, frameBottom, frameRight, frameBottom));
            pageLayout.BorderLines.Add(new GridLine(frameLeft, frameTop, frameLeft, frameBottom));
            pageLayout.BorderLines.Add(new GridLine(frameRight, frameTop, frameRight, frameBottom));

            if (borderOnly) return;

            // --- 2. Text-spill mask ---------------------------------------------------------------

            // spillBlocked[ri, gi] = true means the vertical line at colX[gi+1] for row ri
            // must be suppressed because cell text from this row spills across that gap.
            // gi ranges 0 .. colCount-2.
            var spillBlocked = new bool[rowCount, colCount - 1];

            for (int ri = 0; ri < rowCount; ri++)
            {
                int row = page.FromRow + ri;
                for (int ci = 0; ci < colCount; ci++)
                {
                    int col = page.FromColumn + ci;
                    var cell = page.Map[row, col];

                    if (cell == null || cell.Hidden || (cell.ContentAligmnet != null && cell.ContentAligmnet.WrapText)) continue;
                    // Merged cells never spill — their text is clipped to the merged region.
                    // The vertical lines around the merge are handled by MergedAcross checks below.
                    if (cell.Merged) continue;

                    double spill = cell.TotalTextLength - cell.ColumnWidth;
                    if (spill <= 0d) continue;

                    var halign = cell.ContentAligmnet.HorizontalAlignment;

                    // Right spill — Left / General alignment.
                    if (halign == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Left ||
                        halign == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.General)
                    {
                        MarkSpillRight(page, ri, ci, colCount, colX, spill, spillBlocked);

                    }

                   // Left spill — Right alignment.
                    else if (halign == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Right)
                    {
                        MarkSpillLeft(page, ri, ci, colX, spill, spillBlocked);
                    }
                    // Both directions — Center alignment, half the excess each way.
                    else if (halign == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Center)
                    {
                        double halfSpill = spill / 2d;
                        MarkSpillRight(page, ri, ci, colCount, colX, halfSpill, spillBlocked);
                        MarkSpillLeft(page, ri, ci, colX, halfSpill, spillBlocked);

                   }
                }
            }

            // --- 3. Vertical gridlines (one per internal column gap) -----------------------------

            for (int gi = 0; gi < colCount - 1; gi++)
            {
                double x = colX[gi + 1];
                int leftCol = page.FromColumn + gi;
                int rightCol = leftCol + 1;

                double? runStart = null;
                double runEnd = 0d;

                for (int ri = 0; ri < rowCount; ri++)
                {
                    int row = page.FromRow + ri;

                    // Hidden rows have zero height — no segment to draw.
                    if (page.RowHeights[ri] == 0d)
                    {
                        FlushVertical(pageLayout, x, ref runStart, runEnd);
                        continue;
                    }

                    var leftCell = page.Map[row, leftCol];
                    var rightCell = page.Map[row, rightCol];

                    bool blocked =
                        spillBlocked[ri, gi] ||
                        MergedAcross(leftCell, leftCol) || // left cell's merge extends right past this gap
                        MergedAcross(rightCell, rightCol, fromLeft: true) || // right cell's merge came from the left
                        HasRightBorder(leftCell) ||
                        HasLeftBorder(rightCell);

                    if (!blocked)
                    {
                        if (runStart == null) runStart = rowY[ri];
                        runEnd = rowY[ri + 1];
                    }
                    else
                    {
                        FlushVertical(pageLayout, x, ref runStart, runEnd);
                    }
                }
                FlushVertical(pageLayout, x, ref runStart, runEnd);
            }

            // --- 4. Horizontal gridlines (one per internal row gap) ------------------------------

            for (int ri = 0; ri < rowCount - 1; ri++)
            {
                double y = rowY[ri + 1];
                int topRow = page.FromRow + ri;
                int bottomRow = topRow + 1;

                double? runStart = null;
                double runEnd = 0d;

                for (int ci = 0; ci < colCount; ci++)
                {
                    int col = page.FromColumn + ci;

                    // Hidden columns have zero width — no segment to draw.
                    if (colX[ci + 1] - colX[ci] == 0d)
                    {
                        FlushHorizontal(pageLayout, y, ref runStart, runEnd);
                        continue;
                    }

                    var topCell = page.Map[topRow, col];
                    var bottomCell = page.Map[bottomRow, col];

                    bool blocked =
                        MergedDown(topCell, topRow) || // top cell's merge extends down past this gap
                        MergedDown(bottomCell, bottomRow, fromAbove: true) || // bottom cell's merge came from above
                        HasBottomBorder(topCell) ||
                        HasTopBorder(bottomCell);

                    if (!blocked)
                    {
                        if (runStart == null) runStart = colX[ci];
                        runEnd = colX[ci + 1];
                    }
                    else
                    {
                        FlushHorizontal(pageLayout, y, ref runStart, runEnd);
                    }
                }
                FlushHorizontal(pageLayout, y, ref runStart, runEnd);
            }
        }

        // ---- Spill helpers ----------------------------------------------------------------------

        /// <summary>
        /// Marks column gaps to the right of column <paramref name="ci"/> as spill-blocked for
        /// row <paramref name="ri"/>, until the spill budget runs out or a non-empty cell stops it.
        /// </summary>
        private static void MarkSpillRight(
            Page page, int ri, int ci, int colCount,
            double[] colX, double spill, bool[,] spillBlocked)
        {
            int row = page.FromRow + ri;
            // Walk gaps gi = ci, ci+1, ... (gap gi lies between column gi and column gi+1).
            for (int gi = ci; gi < colCount - 1; gi++)
            {
                // The cell immediately to the right of this gap stops the spill if it has text.
                var rightCell = page.Map[row, page.FromColumn + gi + 1];
                if (rightCell != null && !string.IsNullOrEmpty(rightCell.Text)) break;

                // Distance from the right edge of the source cell to this gap.
                double distToGap = colX[gi + 1] - colX[ci + 1];
                if (spill > distToGap)
                    spillBlocked[ri, gi] = true;
                else
                    break;
            }
        }

        /// <summary>
        /// Marks column gaps to the left of column <paramref name="ci"/> as spill-blocked for
        /// row <paramref name="ri"/>, until the spill budget runs out or a non-empty cell stops it.
        /// </summary>
        private static void MarkSpillLeft(
            Page page, int ri, int ci,
            double[] colX, double spill, bool[,] spillBlocked)
        {
            int row = page.FromRow + ri;
            // Walk gaps gi = ci-1, ci-2, ... (gap gi lies between column gi and column gi+1).
            for (int gi = ci - 1; gi >= 0; gi--)
            {
                // The cell immediately to the left of this gap stops the spill if it has text.
                var leftCell = page.Map[row, page.FromColumn + gi];
                if (leftCell != null && !string.IsNullOrEmpty(leftCell.Text)) break;

                // Distance from the left edge of the source cell to this gap.
                double distToGap = colX[ci] - colX[gi + 1];
                if (spill > distToGap)
                    spillBlocked[ri, gi] = true;
                else
                    break;
            }
        }

        // ---- Run-flush helpers ------------------------------------------------------------------

        private static void FlushVertical(PdfPageLayout layout, double x, ref double? runStart, double runEnd)
        {
            if (runStart == null) return;
            layout.GridLines.Add(new GridLine(x, runStart.Value, x, runEnd));
            runStart = null;
        }

        private static void FlushHorizontal(PdfPageLayout layout, double y, ref double? runStart, double runEnd)
        {
            if (runStart == null) return;
            layout.GridLines.Add(new GridLine(runStart.Value, y, runEnd, y));
            runStart = null;
        }

        // ---- Merge-span helpers -----------------------------------------------------------------

        /// <summary>
        /// Returns true when <paramref name="cell"/> is part of a merged region that spans across
        /// the vertical gap adjacent to <paramref name="col"/>.
        /// <para>
        ///   <paramref name="fromLeft"/> = false (default): cell is the LEFT neighbour of the gap —
        ///   block if the merge extends to the right (toCol > col).
        /// </para>
        /// <para>
        ///   <paramref name="fromLeft"/> = true: cell is the RIGHT neighbour of the gap —
        ///   block if the merge started to the left (fromCol &lt; col).
        /// </para>
        /// </summary>
        private static bool MergedAcross(PdfCell cell, int col, bool fromLeft = false)
        {
            if (cell == null || !cell.Merged || cell.MergedAddress == null) return false;
            return fromLeft
                ? cell.MergedAddress._fromCol < col
                : cell.MergedAddress._toCol > col;
        }

        /// <summary>
        /// Returns true when <paramref name="cell"/> is part of a merged region that spans across
        /// the horizontal gap adjacent to <paramref name="row"/>.
        /// <para>
        ///   <paramref name="fromAbove"/> = false (default): cell is the TOP neighbour of the gap —
        ///   block if the merge extends downward (toRow > row).
        /// </para>
        /// <para>
        ///   <paramref name="fromAbove"/> = true: cell is the BOTTOM neighbour of the gap —
        ///   block if the merge started above (fromRow &lt; row).
        /// </para>
        /// </summary>
        private static bool MergedDown(PdfCell cell, int row, bool fromAbove = false)
        {
            if (cell == null || !cell.Merged || cell.MergedAddress == null) return false;
            return fromAbove
                ? cell.MergedAddress._fromRow < row
                : cell.MergedAddress._toRow > row;
        }

        // ---- Border presence helpers ------------------------------------------------------------
        // For merged cells, use the main cell's style (which holds the region's outer borders).

        private static bool HasRightBorder(PdfCell cell)
        {
            var cs = (cell?.Merged == true && cell.Main != null) ? cell.Main.CellStyle : cell?.CellStyle;
            if (cs == null) return false;
            return (cs.xfRight != null && cs.xfRight.Style != ExcelBorderStyle.None) ||
                   (cs.dxfRight?.HasValue ?? false);

       }

        private static bool HasLeftBorder(PdfCell cell)
        {
            var cs = (cell?.Merged == true && cell.Main != null) ? cell.Main.CellStyle : cell?.CellStyle;
            if (cs == null) return false;
            return (cs.xfLeft != null && cs.xfLeft.Style != ExcelBorderStyle.None) ||
                   (cs.dxfLeft?.HasValue ?? false);
        }

        private static bool HasTopBorder(PdfCell cell)
        {
            var cs = (cell?.Merged == true && cell.Main != null) ? cell.Main.CellStyle : cell?.CellStyle;
            if (cs == null) return false;
            return (cs.xfTop != null && cs.xfTop.Style != ExcelBorderStyle.None) ||
                   (cs.dxfTop?.HasValue ?? false);
        }

        private static bool HasBottomBorder(PdfCell cell)
        {
            var cs = (cell?.Merged == true && cell.Main != null) ? cell.Main.CellStyle : cell?.CellStyle;
            if (cs == null) return false;
            return (cs.xfBottom != null && cs.xfBottom.Style != ExcelBorderStyle.None) ||
                   (cs.dxfBottom?.HasValue ?? false);
        }

    }
}
