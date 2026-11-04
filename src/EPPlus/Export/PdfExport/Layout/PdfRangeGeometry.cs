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

namespace OfficeOpenXml.Export.PdfExport.Layout
{
    internal static class PdfRangeGeometry
    {
        internal static PdfCell RangeCell(PdfRange range, int row, int col)
        {
            if (row < range.Map.FromRow || row > range.Map.ToRow)
            {
                return null;
            }
            if (col < range.Map.FromColumn || col > range.Map.ToColumn)
            {
                return null;
            }
            return range.Map[row, col];
        }

        internal static double RangeColWidth(PdfRange range, int representativeRow, int col)
        {
            var cell = RangeCell(range, representativeRow, col);
            return cell?.ColumnWidth ?? 0d;
        }

        internal static double RangeRowHeight(PdfRange range, int row)
        {
            int i = row - range.Range._fromRow;
            return (i >= 0 && i < range.RowHeights.Count) ? range.RowHeights[i].Height : 0d;
        }
    }
}
