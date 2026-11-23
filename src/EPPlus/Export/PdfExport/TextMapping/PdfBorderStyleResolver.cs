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
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Table;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Export.PdfExport.TextMapping
{
    internal static class PdfBorderStyleResolver
    {
        internal static class TableEdgeOrder
        {
            public const int None = 0, WholeTable = 1,
                FirstColumnStripe = 2, SecondColumnStripe = 3,
                FirstRowStripe = 4, SecondRowStripe = 5,
                LastColumn = 6, FirstColumn = 7,
                TotalRow = 8, HeaderRow = 9,
                FirstHeaderCell = 10, LastHeaderCell = 11,
                FirstTotalCell = 12, LastTotalCell = 13,
                ConditionalFormat = 50,
                UserSet = 100;
        }

        internal static void GetBorderStyles(ExcelRangeBase cell, PdfCellStyle cellStyle, PdfCell pcell, Dictionary<ExcelTable, ExcelTableNamedStyle> tableStyleCache)
        {
            if (cell != null)
            {
                cellStyle.xfTop = cell.Style.Border.Top;
                cellStyle.xfBottom = cell.Style.Border.Bottom;
                cellStyle.xfLeft = cell.Style.Border.Left;
                cellStyle.xfRight = cell.Style.Border.Right;
                if (pcell.Main == null)
                {
                    cellStyle.Diagonal = cell.Style.Border.Diagonal;
                    cellStyle.DiagonalUp = cell.Style.Border.DiagonalUp;
                    cellStyle.DiagonalDown = cell.Style.Border.DiagonalDown;
                }
                else
                {
                    cellStyle.Diagonal = cell.Style.Border.Diagonal;
                    cellStyle.DiagonalUp = false;
                    cellStyle.DiagonalDown = false;
                }
                var tables = cell.Worksheet.Tables.GetIntersectingRanges(cell);
                if (tables.Count > 0)
                {
                    var table = tables[0].Value;
                    ExcelTableNamedStyle tableStyle = PdfTableStyleLookup.GetTableStyle(table, tableStyleCache);
                    if (tableStyle != null)
                    {
                        cellStyle.dxfTop = GetTopBorderItem(cell, cellStyle.xfTop, table, tableStyle, out int topOrder);
                        cellStyle.dxfTopElementOrder = topOrder;
                        cellStyle.dxfBottom = GetBottomBorderItem(cell, cellStyle.xfBottom, table, tableStyle, out int botOrder);
                        cellStyle.dxfBottomElementOrder = botOrder;
                        cellStyle.dxfLeft = GetLeftBorderItem(cell, cellStyle.xfLeft, table, tableStyle, out int leftOrder);
                        cellStyle.dxfLeftElementOrder = leftOrder;
                        cellStyle.dxfRight = GetRightBorderItem(cell, cellStyle.xfRight, table, tableStyle, out int rightOrder);
                        cellStyle.dxfRightElementOrder = rightOrder;
                    }
                }
                var cfBorder = GetConditionalFormattingBorder(cell);
                if (cfBorder != null)
                {
                    if (cfBorder.Top != null && cfBorder.Top.HasValue)
                    {
                        cellStyle.dxfTop = cfBorder.Top;
                        cellStyle.dxfTopElementOrder = TableEdgeOrder.ConditionalFormat;
                    }
                    if (cfBorder.Bottom != null && cfBorder.Bottom.HasValue)
                    {
                        cellStyle.dxfBottom = cfBorder.Bottom;
                        cellStyle.dxfBottomElementOrder = TableEdgeOrder.ConditionalFormat;
                    }
                    if (cfBorder.Left != null && cfBorder.Left.HasValue)
                    {
                        cellStyle.dxfLeft = cfBorder.Left;
                        cellStyle.dxfLeftElementOrder = TableEdgeOrder.ConditionalFormat;
                    }
                    if (cfBorder.Right != null && cfBorder.Right.HasValue)
                    {
                        cellStyle.dxfRight = cfBorder.Right;
                        cellStyle.dxfRightElementOrder = TableEdgeOrder.ConditionalFormat;
                    }
                }
            }
        }

        internal static ExcelDxfBorderBase GetConditionalFormattingBorder(ExcelRangeBase cell)
        {
            var cf = cell.ConditionalFormatting.GetConditionalFormattings();
            if (cf != null && cf.Count > 0)
            {
                var ordered = cf.OrderBy(r => r.Priority);
                foreach (var rule in ordered)
                {
                    if (!rule.ShouldApplyToCell(cell))
                    {
                        if (rule.StopIfTrue) break;
                        continue;
                    }
                    if (rule.Style?.Border != null && rule.Style.Border.HasValue)
                    {
                        return rule.Style.Border;
                    }
                    if (rule.StopIfTrue) break;
                }
            }
            return null;
        }

        internal static ExcelDxfBorderItem GetTopBorderItem(ExcelRangeBase cell, ExcelBorderItem xfBorder, ExcelTable table, ExcelTableNamedStyle tableStyle, out int elementOrder)
        {
            var range = table.Range;
            int tableRow = cell._fromRow - range._fromRow;
            int tableCol = cell._fromCol - range._fromCol;
            int ts = table.ShowHeader ? 1 : 0;
            var top = tableRow == 0 ? tableStyle.WholeTable.Style.Border.Top : tableStyle.WholeTable.Style.Border.Horizontal;
            elementOrder = TableEdgeOrder.WholeTable;
            if (table.ShowHeader && tableRow == 0)
            {
                if (tableStyle.HeaderRow.Style.Border.Top.HasValue)
                { top = tableStyle.HeaderRow.Style.Border.Top; elementOrder = TableEdgeOrder.HeaderRow; }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstHeaderCell.Style.Border.Top.HasValue)
                { top = tableStyle.FirstHeaderCell.Style.Border.Top; elementOrder = TableEdgeOrder.FirstHeaderCell; }
                if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastHeaderCell.Style.Border.Top.HasValue)
                { top = tableStyle.LastHeaderCell.Style.Border.Top; elementOrder = TableEdgeOrder.LastHeaderCell; }
            }
            else if (table.ShowTotal && cell._fromRow == range._toRow)
            {
                if (tableStyle.TotalRow.Style.Border.Top.HasValue)
                { top = tableStyle.TotalRow.Style.Border.Top; elementOrder = TableEdgeOrder.TotalRow; }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstTotalCell.Style.Border.Top.HasValue)
                { top = tableStyle.FirstTotalCell.Style.Border.Top; elementOrder = TableEdgeOrder.FirstTotalCell; }
                if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastTotalCell.Style.Border.Top.HasValue)
                { top = tableStyle.LastTotalCell.Style.Border.Top; elementOrder = TableEdgeOrder.LastTotalCell; }
            }
            else
            {
                if (table.ShowColumnStripes && (tableCol & 1) == 0)
                {
                    var band = tableStyle.FirstColumnStripe.Style.Border;
                    if (tableRow == 0)
                    { if (band.Top.HasValue) { top = band.Top; elementOrder = TableEdgeOrder.FirstColumnStripe; } }
                    else if (band.Horizontal.HasValue)
                    { top = band.Horizontal; elementOrder = TableEdgeOrder.FirstColumnStripe; }
                }
                if (table.ShowColumnStripes && (tableCol & 1) != 0)
                {
                    var band = tableStyle.SecondColumnStripe.Style.Border;
                    if (tableRow == 0)
                    { if (band.Top.HasValue) { top = band.Top; elementOrder = TableEdgeOrder.SecondColumnStripe; } }
                    else if (band.Horizontal.HasValue)
                    { top = band.Horizontal; elementOrder = TableEdgeOrder.SecondColumnStripe; }
                }
                if (table.ShowRowStripes && tableStyle.FirstRowStripe.Style.Border.Top.HasValue && (tableRow & 1) != 0)
                { top = tableStyle.FirstRowStripe.Style.Border.Top; elementOrder = TableEdgeOrder.FirstRowStripe; }
                if (table.ShowRowStripes && tableStyle.SecondRowStripe.Style.Border.Top.HasValue && (tableRow & 1) == 0)
                { top = tableStyle.SecondRowStripe.Style.Border.Top; elementOrder = TableEdgeOrder.SecondRowStripe; }
                if (table.ShowLastColumn && tableStyle.LastColumn.Style.Border.Top.HasValue && cell._fromCol == range._toCol)
                { top = tableStyle.LastColumn.Style.Border.Top; elementOrder = TableEdgeOrder.LastColumn; }
                if (table.ShowFirstColumn && tableStyle.FirstColumn.Style.Border.Top.HasValue && tableCol == 0)
                { top = tableStyle.FirstColumn.Style.Border.Top; elementOrder = TableEdgeOrder.FirstColumn; }
            }
            if (top == null || !top.Style.HasValue || top.Style.Value == ExcelBorderStyle.None) elementOrder = TableEdgeOrder.None;
            return top;
        }

        internal static ExcelDxfBorderItem GetBottomBorderItem(ExcelRangeBase cell, ExcelBorderItem xfBorder, ExcelTable table, ExcelTableNamedStyle tableStyle, out int elementOrder)
        {
            var range = table.Range;
            int tableRow = cell._fromRow - range._fromRow;
            int tableCol = cell._fromCol - range._fromCol;
            var bottom = range._toRow == cell._fromRow ? tableStyle.WholeTable.Style.Border.Bottom : null;
            elementOrder = TableEdgeOrder.WholeTable;
            if (table.ShowHeader && tableRow == 0)
            {
                if (tableStyle.HeaderRow.Style.Border.Bottom.HasValue)
                { bottom = tableStyle.HeaderRow.Style.Border.Bottom; elementOrder = TableEdgeOrder.HeaderRow; }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstHeaderCell.Style.Border.Bottom.HasValue)
                { bottom = tableStyle.FirstHeaderCell.Style.Border.Bottom; elementOrder = TableEdgeOrder.FirstHeaderCell; }
                if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastHeaderCell.Style.Border.Bottom.HasValue)
                { bottom = tableStyle.LastHeaderCell.Style.Border.Bottom; elementOrder = TableEdgeOrder.LastHeaderCell; }
            }
            else if (table.ShowTotal && cell._fromRow == range._toRow)
            {
                if (tableStyle.TotalRow.Style.Border.Bottom.HasValue)
                { bottom = tableStyle.TotalRow.Style.Border.Bottom; elementOrder = TableEdgeOrder.TotalRow; }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstTotalCell.Style.Border.Bottom.HasValue)
                { bottom = tableStyle.FirstTotalCell.Style.Border.Bottom; elementOrder = TableEdgeOrder.FirstTotalCell; }
                if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastTotalCell.Style.Border.Bottom.HasValue)
                { bottom = tableStyle.LastTotalCell.Style.Border.Bottom; elementOrder = TableEdgeOrder.LastTotalCell; }
            }
            else
            {
                if (table.ShowColumnStripes && tableStyle.FirstColumnStripe.Style.Border.Bottom.HasValue && (tableCol & 1) != 0)
                {
                    if (cell._fromRow > range._fromRow && cell._fromRow < range._toRow)
                    { bottom = tableStyle.FirstColumnStripe.Style.Border.Horizontal; elementOrder = TableEdgeOrder.FirstColumnStripe; }
                    else if (cell._fromRow < range._toRow)
                    { bottom = null; }
                    else
                    { bottom = tableStyle.FirstColumnStripe.Style.Border.Bottom; elementOrder = TableEdgeOrder.FirstColumnStripe; }
                }
                if (table.ShowColumnStripes && tableStyle.SecondColumnStripe.Style.Border.Bottom.HasValue && (tableCol & 1) == 0)
                {
                    if (cell._fromRow > range._fromRow && cell._fromRow < range._toRow)
                    { bottom = tableStyle.SecondColumnStripe.Style.Border.Horizontal; elementOrder = TableEdgeOrder.SecondColumnStripe; }
                    else if (cell._fromRow < range._toRow)
                    { bottom = null; }
                    else
                    { bottom = tableStyle.SecondColumnStripe.Style.Border.Bottom; elementOrder = TableEdgeOrder.SecondColumnStripe; }
                }
                if (table.ShowRowStripes && tableStyle.FirstRowStripe.Style.Border.Bottom.HasValue && (tableRow & 1) != 0)
                { bottom = tableStyle.FirstRowStripe.Style.Border.Bottom; elementOrder = TableEdgeOrder.FirstRowStripe; }
                if (table.ShowRowStripes && tableStyle.SecondRowStripe.Style.Border.Bottom.HasValue && (tableRow & 1) == 0)
                { bottom = tableStyle.SecondRowStripe.Style.Border.Bottom; elementOrder = TableEdgeOrder.SecondRowStripe; }
                if (table.ShowLastColumn && tableStyle.LastColumn.Style.Border.Bottom.HasValue && cell._fromCol == range._toCol)
                { bottom = tableStyle.LastColumn.Style.Border.Bottom; elementOrder = TableEdgeOrder.LastColumn; }
                if (table.ShowFirstColumn && tableStyle.FirstColumn.Style.Border.Bottom.HasValue && tableCol == 0)
                { bottom = tableStyle.FirstColumn.Style.Border.Bottom; elementOrder = TableEdgeOrder.FirstColumn; }
            }
            if (bottom == null || !bottom.Style.HasValue || bottom.Style.Value == ExcelBorderStyle.None) elementOrder = TableEdgeOrder.None;
            return bottom;
        }

        internal static ExcelDxfBorderItem GetLeftBorderItem(ExcelRangeBase cell, ExcelBorderItem xfBorder, ExcelTable table, ExcelTableNamedStyle tableStyle, out int elementOrder)
        {
            var range = table.Range;
            int tableRow = cell._fromRow - range._fromRow;
            int tableCol = cell._fromCol - range._fromCol;
            var left = tableCol == 0 ? tableStyle.WholeTable.Style.Border.Left : tableStyle.WholeTable.Style.Border.Vertical;
            elementOrder = TableEdgeOrder.WholeTable;
            if (table.ShowHeader && tableRow == 0)
            {
                if (tableStyle.HeaderRow.Style.Border.Left.HasValue)
                { left = tableStyle.HeaderRow.Style.Border.Left; elementOrder = TableEdgeOrder.HeaderRow; }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstHeaderCell.Style.Border.Left.HasValue)
                { left = tableStyle.FirstHeaderCell.Style.Border.Left; elementOrder = TableEdgeOrder.FirstHeaderCell; }
                if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastHeaderCell.Style.Border.Left.HasValue)
                { left = tableStyle.LastHeaderCell.Style.Border.Left; elementOrder = TableEdgeOrder.LastHeaderCell; }
            }
            else if (table.ShowTotal && cell._fromRow == range._toRow)
            {
                if (tableStyle.TotalRow.Style.Border.Left.HasValue)
                { left = tableStyle.TotalRow.Style.Border.Left; elementOrder = TableEdgeOrder.TotalRow; }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstTotalCell.Style.Border.Left.HasValue)
                { left = tableStyle.FirstTotalCell.Style.Border.Left; elementOrder = TableEdgeOrder.FirstTotalCell; }
                if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastTotalCell.Style.Border.Left.HasValue)
                { left = tableStyle.LastTotalCell.Style.Border.Left; elementOrder = TableEdgeOrder.LastTotalCell; }
            }
            else
            {
                if (table.ShowColumnStripes && tableStyle.FirstColumnStripe.Style.Border.Left.HasValue && (tableCol & 1) != 0)
                { left = tableStyle.FirstColumnStripe.Style.Border.Left; elementOrder = TableEdgeOrder.FirstColumnStripe; }
                if (table.ShowColumnStripes && tableStyle.SecondColumnStripe.Style.Border.Left.HasValue && (tableCol & 1) == 0)
                { left = tableStyle.SecondColumnStripe.Style.Border.Left; elementOrder = TableEdgeOrder.SecondColumnStripe; }
                if (table.ShowRowStripes && tableStyle.FirstRowStripe.Style.Border.Left.HasValue && (tableRow & 1) != 0)
                {
                    if (cell._fromCol > range._fromCol)
                    { left = tableStyle.FirstRowStripe.Style.Border.Vertical; elementOrder = TableEdgeOrder.FirstRowStripe; }
                    else
                    { left = tableStyle.FirstRowStripe.Style.Border.Left; elementOrder = TableEdgeOrder.FirstRowStripe; }
                }
                if (table.ShowRowStripes && tableStyle.SecondRowStripe.Style.Border.Left.HasValue && (tableRow & 1) == 0)
                {
                    if (cell._fromCol > range._fromCol)
                    { left = tableStyle.SecondRowStripe.Style.Border.Vertical; elementOrder = TableEdgeOrder.SecondRowStripe; }
                    else
                    { left = tableStyle.SecondRowStripe.Style.Border.Left; elementOrder = TableEdgeOrder.SecondRowStripe; }
                }
                if (table.ShowLastColumn && tableStyle.LastColumn.Style.Border.Left.HasValue && cell._fromCol == range._toCol)
                { left = tableStyle.LastColumn.Style.Border.Left; elementOrder = TableEdgeOrder.LastColumn; }
                if (table.ShowFirstColumn && tableStyle.FirstColumn.Style.Border.Left.HasValue && tableCol == 0)
                { left = tableStyle.FirstColumn.Style.Border.Left; elementOrder = TableEdgeOrder.FirstColumn; }
            }
            if (left == null || !left.Style.HasValue || left.Style.Value == ExcelBorderStyle.None) elementOrder = TableEdgeOrder.None;
            return left;
        }

        internal static ExcelDxfBorderItem GetRightBorderItem(ExcelRangeBase cell, ExcelBorderItem xfBorder, ExcelTable table, ExcelTableNamedStyle tableStyle, out int elementOrder)
        {
            var range = table.Range;
            int tableRow = cell._fromRow - range._fromRow;
            int tableCol = cell._fromCol - range._fromCol;
            var right = cell._fromCol == range._toCol ? tableStyle.WholeTable.Style.Border.Right : tableStyle.WholeTable.Style.Border.Vertical;
            elementOrder = TableEdgeOrder.WholeTable;
            if (table.ShowHeader && tableRow == 0)
            {
                if (tableStyle.HeaderRow.Style.Border.Right.HasValue)
                { right = tableStyle.HeaderRow.Style.Border.Right; elementOrder = TableEdgeOrder.HeaderRow; }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstHeaderCell.Style.Border.Right.HasValue)
                { right = tableStyle.FirstHeaderCell.Style.Border.Right; elementOrder = TableEdgeOrder.FirstHeaderCell; }
                if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastHeaderCell.Style.Border.Right.HasValue)
                { right = tableStyle.LastHeaderCell.Style.Border.Right; elementOrder = TableEdgeOrder.LastHeaderCell; }
            }
            else if (table.ShowTotal && cell._fromRow == range._toRow)
            {
                if (tableStyle.TotalRow.Style.Border.Right.HasValue)
                { right = tableStyle.TotalRow.Style.Border.Right; elementOrder = TableEdgeOrder.TotalRow; }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstTotalCell.Style.Border.Right.HasValue)
                { right = tableStyle.FirstTotalCell.Style.Border.Right; elementOrder = TableEdgeOrder.FirstTotalCell; }
                if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastTotalCell.Style.Border.Right.HasValue)
                { right = tableStyle.LastTotalCell.Style.Border.Right; elementOrder = TableEdgeOrder.LastTotalCell; }
            }
            else
            {
                if (table.ShowColumnStripes && tableStyle.FirstColumnStripe.Style.Border.Right.HasValue && (tableCol & 1) != 0)
                { right = tableStyle.FirstColumnStripe.Style.Border.Right; elementOrder = TableEdgeOrder.FirstColumnStripe; }
                if (table.ShowColumnStripes && tableStyle.SecondColumnStripe.Style.Border.Right.HasValue && (tableCol & 1) == 0)
                { right = tableStyle.SecondColumnStripe.Style.Border.Right; elementOrder = TableEdgeOrder.SecondColumnStripe; }
                if (table.ShowRowStripes && tableStyle.FirstRowStripe.Style.Border.Right.HasValue && (tableRow & 1) != 0)
                {
                    if (cell._fromCol < range._toCol)
                    { right = tableStyle.FirstRowStripe.Style.Border.Vertical; elementOrder = TableEdgeOrder.FirstRowStripe; }
                    else
                    { right = tableStyle.FirstRowStripe.Style.Border.Right; elementOrder = TableEdgeOrder.FirstRowStripe; }
                }
                if (table.ShowRowStripes && tableStyle.SecondRowStripe.Style.Border.Right.HasValue && (tableRow & 1) == 0)
                {
                    if (cell._fromCol < range._toCol)
                    { right = tableStyle.SecondRowStripe.Style.Border.Vertical; elementOrder = TableEdgeOrder.SecondRowStripe; }
                    else
                    { right = tableStyle.SecondRowStripe.Style.Border.Right; elementOrder = TableEdgeOrder.SecondRowStripe; }
                }
                if (table.ShowLastColumn && tableStyle.LastColumn.Style.Border.Right.HasValue && cell._fromCol == range._toCol)
                { right = tableStyle.LastColumn.Style.Border.Right; elementOrder = TableEdgeOrder.LastColumn; }
                if (table.ShowFirstColumn && tableStyle.FirstColumn.Style.Border.Right.HasValue && tableCol == 0)
                { right = tableStyle.FirstColumn.Style.Border.Right; elementOrder = TableEdgeOrder.FirstColumn; }
            }
            if (right == null || !right.Style.HasValue || right.Style.Value == ExcelBorderStyle.None) elementOrder = TableEdgeOrder.None;
            return right;
        }
    }
}
