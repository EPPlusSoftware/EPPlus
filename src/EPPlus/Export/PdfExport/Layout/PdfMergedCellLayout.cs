using EPPlus.Export.Pdf.Layout;
using EPPlus.Export.Pdf.Settings;
using OfficeOpenXml.Export.PdfExport.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.PdfExport.Layout
{
    internal class PdfMergedCellLayout
    {
        internal static Pages PrecomputeMergedCells(PdfPageSettings pageSettings, PdfRange range, Pages pdfPages)
        {
            for (int i = 0; i < pdfPages.Page.Length; i++)
                pdfPages.Page[i] = PrecomputePageMergedCells(pageSettings, range, pdfPages.Page[i]);
            return pdfPages;
        }

        private static Page PrecomputePageMergedCells(PdfPageSettings pageSettings, PdfRange range, Page page)
        {
            page.MergedCells = new Dictionary<string, MergedCellDrawInfo>();
            var colX = PdfLayout.BuildColumnXPositions(pageSettings, page);
            for (int row = page.FromRow; row <= page.ToRow; row++)
            {
                for (int col = page.FromColumn; col <= page.ToColumn; col++)
                {
                    var cell = page.Map[row, col];
                    if (cell == null || !cell.Merged) continue;
                    string key = cell.MergedAddress.Address;
                    if (page.MergedCells.ContainsKey(key)) continue;
                    var addr = cell.MergedAddress;
                    double drawX = colX[col - page.FromColumn];
                    for (int c = addr._fromCol; c < col; c++)
                    {
                        int rangeIdx = c - range.Range._fromCol;
                        if (rangeIdx >= 0 && rangeIdx < range.ColWidths.Count)
                            drawX -= range.ColWidths[rangeIdx];
                    }
                    double drawY = PdfLayout.GetOriginY(pageSettings, page) - page.HeadingHeight - page.PrintTitleHeight;
                    for (int r = page.FromRow; r < row; r++)
                    {
                        drawY -= range.RowHeights[r - range.Range._fromRow].Height;
                    }
                    for (int r = addr._fromRow; r < row; r++)
                    {
                        int rangeIdx = r - range.Range._fromRow;
                        if (rangeIdx >= 0 && rangeIdx < range.RowHeights.Count)
                            drawY += range.RowHeights[rangeIdx].Height;
                    }
                    double mergeWidth = 0d;
                    for (int c = addr._fromCol; c <= addr._toCol; c++)
                    {
                        int rangeIdx = c - range.Range._fromCol;
                        if (rangeIdx >= 0 && rangeIdx < range.ColWidths.Count)
                            mergeWidth += range.ColWidths[rangeIdx];
                    }
                    double mergeHeight = 0d;
                    for (int r = addr._fromRow; r <= addr._toRow; r++)
                    {
                        int rangeIdx = r - range.Range._fromRow;
                        if (rangeIdx >= 0 && rangeIdx < range.RowHeights.Count)
                            mergeHeight += range.RowHeights[rangeIdx].Height;
                    }
                    page.MergedCells[key] = new MergedCellDrawInfo
                    {
                        X = drawX,
                        Y = drawY,
                        Width = mergeWidth,
                        Height = mergeHeight
                    };
                }
            }
            return page;
        }
    }
}
