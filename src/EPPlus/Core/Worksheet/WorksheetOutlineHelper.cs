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
        internal int CollapseRow(int rowNo, int level, bool collapsed, bool collapseChildren, int addValue, bool parentIsHidden = false)
        {
            var row = GetRow(rowNo);
            int startLevel = 0;
            if (row != null)
            {
                startLevel = row.OutlineLevel;
            }
            if (row==null)
            {
                _worksheet.Row(rowNo).Collapsed=collapsed;

            }
            else
            {
                row.Collapsed = collapsed;
            }

            bool? hidden;
            if (collapsed)
            {
                hidden = null;
            }
            else
            {
                //Check if the parent row is hidden.
                hidden = row.Hidden;
            }

            var r = rowNo + addValue;
            row = GetRow(r);
            while (row != null && (row.OutlineLevel > startLevel || (row.OutlineLevel >= level && level>=0)))
            {
                if (level < row.OutlineLevel)
                {
                    row.Hidden = hidden ?? collapsed;
                    if (collapseChildren && level != -2) row.Collapsed = collapsed;
                }
                else
                {
                    if (collapseChildren)
                    {
                        row.Collapsed = true;
                    }
                    else
                    {
                        row.Hidden = hidden ?? !collapsed;
                        if (level > row.OutlineLevel) row.Collapsed = !collapsed;
                    }
                }
                
                if (addValue < 0)
                {
                    row = GetRow(r--);
                }
                else
                {
                    row = GetRow(r++);
                }

                if (row != null) rowNo = r;
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
            if (col == null)
            {
                col = _worksheet.Column(colNo);
            }

            col.Collapsed = collapsed;

            bool? hidden;
            if (collapsed)
            {
                hidden = null;
            }
            else
            {
                //Check if the parent row is hidden.
                hidden = col.Hidden;
            }

            col = GetColumn(colNo + addValue);
            while(col!=null && (col.OutlineLevel > startLevel || (col.OutlineLevel >= level && level >= 0)))
            {
                if (level < col.OutlineLevel)
                {
                    col.Hidden = hidden??collapsed;
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
                        col.Hidden = hidden??!collapsed;
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
