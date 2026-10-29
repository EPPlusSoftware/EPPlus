using OfficeOpenXml.Export.PdfExport.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.PdfExport.Layout
{
    internal class PdfRangeGeometry
    {
        internal static PdfCell RangeCell(PdfRange range, int row, int col)
        {
            if (row < range.Map.FromRow || row > range.Map.ToRow) return null;
            if (col < range.Map.FromColumn || col > range.Map.ToColumn) return null;
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
