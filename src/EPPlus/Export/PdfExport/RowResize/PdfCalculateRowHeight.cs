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
using OfficeOpenXml.Export.PdfExport.Data;

namespace OfficeOpenXml.Export.PdfExport.RowResize
{
    internal class PdfCalculateRowHeight
    {
        public static void ResizeRowHeights(PdfWorksheet pdfSheet)
        {
            for (int r = 0; r < pdfSheet.Ranges.Count; r++)
            {
                var range = pdfSheet.Ranges[r];
                ResizeRange(ref range);
                pdfSheet.Ranges[r] = range;
            }
            if (pdfSheet.CommentsAndNotes.Map != null)
            {
                var cnRange = pdfSheet.CommentsAndNotes;
                ResizeRange(ref cnRange);
                pdfSheet.CommentsAndNotes = cnRange;
            }
        }

        public static void ResizeRange(ref PdfRange range)
        {
            double newTotalHeight = 0d;
            for (int rowIdx = 0; rowIdx < range.RowHeights.Count; rowIdx++)
            {
                var rowHeight = range.RowHeights[rowIdx];
                if (rowHeight.Height == 0d)
                    continue;

                if (!rowHeight.UsesDefaultValue)
                {
                    newTotalHeight += rowHeight.Height;
                    continue;
                }
                int row = range.Range._fromRow + rowIdx;
                double maxRequired = rowHeight.Height;
                bool grew = false;
                for (int colIdx = 0; colIdx < range.ColWidths.Count; colIdx++)
                {
                    int col = range.Range._fromCol + colIdx;
                    var cell = range.Map[row, col];
                    if (cell == null || cell.Hidden)
                        continue;
                    if (cell.TextLines == null || cell.TextLines.Count == 0)
                        continue;
                    if (cell.ContentAligmnet == null || cell.ContentAligmnet.ShrinkToFit)
                        continue;

                    double required;
                    if (cell.Merged)
                    {
                        if (cell.MergedAddress == null || cell.MergedAddress.Start.Row != cell.MergedAddress.End.Row)
                            continue;
                        required = GetMaxLineHeight(cell);
                    }
                    else
                    {
                        required = GetRequiredHeightFromLines(cell);
                    }
                    if (required > maxRequired)
                    {
                        maxRequired = required;
                        grew = true;
                    }
                }
                if (grew)
                {
                    rowHeight.Height = maxRequired;
                    range.RowHeights[rowIdx] = rowHeight;
                }
                newTotalHeight += rowHeight.Height;
            }
            range.TotalHeight = newTotalHeight;
        }

        private static double GetRequiredHeightFromLines(PdfCell cell)
        {
            double total = 0d;
            foreach (var line in cell.TextLines)
            {
                total += line.LargestAscent + line.LargestDescent;
            }
            return total;
        }

        private static double GetMaxLineHeight(PdfCell cell)
        {
            double max = 0d;
            foreach (var line in cell.TextLines)
            {
                double h = line.LargestAscent + line.LargestDescent;
                if (h > max) max = h;
            }
            return max;
        }
    }
}
