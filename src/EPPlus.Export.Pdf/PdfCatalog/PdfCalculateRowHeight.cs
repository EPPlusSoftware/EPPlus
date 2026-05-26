using EPPlus.Graphics.Units;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfCatalog
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
                bool hasWrappedCell = false;

                for (int colIdx = 0; colIdx < range.ColWidths.Count; colIdx++)
                {
                    int col = range.Range._fromCol + colIdx;
                    var cell = range.Map[row, col];

                    if (cell == null || cell.Hidden)
                        continue;
                    if (cell.Merged)
                        continue;
                    if (cell.ContentAligmnet.ShrinkToFit)
                        continue;
                    if (!cell.ContentAligmnet.WrapText)
                        continue;
                    if (cell.TextLines == null || cell.TextLines.Count == 0)
                        continue;

                    hasWrappedCell = true;
                    double required = GetRequiredHeightFromLines(cell);
                    if (required > maxRequired)
                        maxRequired = required;
                }

                if (hasWrappedCell)
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
    }
}
