using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet
{
    internal class WorksheetOutlineHelper
    {
        ExcelWorksheet _worksheet;
        internal WorksheetOutlineHelper(ExcelWorksheet worksheet)
        {
            _worksheet = worksheet;
        }
        #region Row
        /// <summary>
        /// Collapses a rows outline from the level
        /// </summary>
        /// <param name="rowNo">The row to collapse</param>
        /// <param name="level">The level to collapse to. -1 means collapse all levels below. -2 Means collapse the row (rowNo) only.</param>
        /// <param name="collapsed"></param>
        /// <returns></returns>
        internal int CollapseRowBottom(int rowNo, int level, bool collapsed)
        {
            int lvl;
            var row = GetRow(rowNo);
            if (level == -2) level = row?.OutlineLevel ?? 0;
            if (row == null || row.OutlineLevel == 0)
            {
                row = GetRow(--rowNo);
                if (row == null || row.OutlineLevel == 0)
                {
                    return rowNo;
                }
                else
                {
                    _worksheet.Row(rowNo + 1).Collapsed = collapsed;
                    lvl = 0;
                    row.Hidden = collapsed;
                }
            }
            else
            {
                row.Collapsed = collapsed;
                lvl = row.OutlineLevel;
            }


            row = GetRow(--rowNo);

            while (row != null && row.OutlineLevel > lvl)
            {
                row.Hidden = collapsed;
                if (level == -1 || level < row.OutlineLevel)
                {
                    row.Collapsed = collapsed;
                }
                row = GetRow(--rowNo);
            }

            return rowNo;
        }

        internal int CollapseRowTop(int rowNo, int level, bool collapsed)
        {
            int lvl;
            var row = GetRow(rowNo);
            if (level == -2) level = row?.OutlineLevel ?? 0;
            if (row == null || row.OutlineLevel == 0)
            {
                row = GetRow(++rowNo);
                if (row == null || row.OutlineLevel == 0 || rowNo > ExcelPackage.MaxRows)
                {
                    return rowNo;
                }
                else
                {
                    _worksheet.Row(rowNo - 1).Collapsed = collapsed;
                    lvl = 0;
                    row.Hidden = collapsed;
                }
            }
            else
            {
                lvl = row.OutlineLevel;
                row.Collapsed = collapsed;
            }

            row = GetRow(++rowNo);
            while (row != null && row.OutlineLevel >= lvl)
            {
                row.Hidden = collapsed;
                if (level == -1 || level <= row.OutlineLevel)
                {
                    row.Collapsed = collapsed;
                }
                row = GetRow(++rowNo);
            }

            return rowNo;
        }
        private RowInternal GetRow(int row)
        {
            if (row < 1 || row > ExcelPackage.MaxRows) return null;
            return _worksheet.GetValueInner(row, 0) as RowInternal;
        }
        #endregion
        #region Column
        internal int CollapseColumn(int colNo, int level, bool collapsed, bool collapseChildren, int addValue)
        {
            var col = GetColumn(colNo);
            int startLevel = 0;
            if(col!=null)
            {
                startLevel = col.OutlineLevel;
            }
            if(level < col?.OutlineLevel) 
                col.Collapsed = collapsed;
            else 
                _worksheet.Column(colNo).Collapsed = !collapsed;

            col = GetColumn(colNo + addValue);
            while(col!=null && col.OutlineLevel > startLevel)
            {
                if (level < col.OutlineLevel)
                {
                    col.Hidden = collapsed;
                    if (collapseChildren && level != -2) col.Collapsed = collapsed;
                }
                else
                {
                    if (collapseChildren)
                    {
                        col.Collapsed = true;
                    }
                    else
                    {
                        col.Hidden = !collapsed;
                        if(level > col.OutlineLevel) col.Collapsed = !collapsed;
                    }
                }
                if(addValue<0)
                {
                    col = GetColumn(col.ColumnMin - 1);
                }
                else
                {
                    col = GetColumn(col.ColumnMax + 1);
                }
                if (col != null) colNo = col.ColumnMax;
            }

            return colNo;
        }
        //internal int CollapseColumnLeft(int colNo, bool allLevels, bool collapsed)
        //{
        //    var col = GetColumn(colNo);
        //    if (col == null || col.OutlineLevel == 0)
        //    {
        //        return colNo;
        //    }

        //    if (col.ColumnMax > colNo) return col.ColumnMax;
        //    var lvl = col.OutlineLevel;
        //    col.Collapsed = collapsed;
        //    col = GetColumn(col.ColumnMax + 1, true);

        //    while (col != null && col.OutlineLevel > lvl)
        //    {
        //        col.Hidden = true;
        //        if (allLevels)
        //        {
        //            col.Collapsed = collapsed;
        //        }
        //        if (allLevels) colNo = col.ColumnMax + 1;
        //        col = _worksheet.GetValueInner(0, col.ColumnMax + 1) as ExcelColumn;
        //    }

        //    return colNo;
        //}
        private ExcelColumn GetColumn(int col, bool ignoreFromCol = true)
        {
            if (col < 1) return null;
            var currentCol = _worksheet.GetValueInner(0, col) as ExcelColumn;
            if (currentCol == null)
            {
                int r = 0, c = col;
                if (_worksheet._values.PrevCell(ref r, ref c))
                {
                    if (c > 0)
                    {
                        ExcelColumn prevCol = _worksheet.GetValueInner(0, c) as ExcelColumn;
                        if (prevCol.ColumnMax < col)
                        {
                            return null;
                        }
                        return prevCol;
                    }
                }
            }
            return currentCol;
        }

        #endregion
    }
}
