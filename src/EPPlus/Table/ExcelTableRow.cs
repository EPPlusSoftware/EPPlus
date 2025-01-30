using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table
{
    /// <summary>
    /// Represents a
    /// </summary>
    public class ExcelTableRow
    {
        internal ExcelTableRow(ExcelTable table, int rowIx)
        {
            _table = table;
            _rowIx = rowIx;
            var startCol = _table.DataRange._fromCol;
            var endCol = _table.DataRange._toCol;
            var startRow = _table.DataRange._fromRow;
            _rowRange = _table.WorkSheet.Cells[_rowIx, startCol, _rowIx, endCol];
        }

        private readonly ExcelTable _table;
        private readonly int _rowIx;
        private readonly ExcelRangeBase _rowRange;

        /// <summary>
        /// Number of columns
        /// </summary>
        public int ColumnCount => _table.Columns.Count;

        /// <summary>
        /// Returns true if the entire row is hidden
        /// </summary>
        public bool IsHidden
        {
            get
            {
                return _table.WorkSheet.Row(_rowIx).Hidden;
            }
        }


        /// <summary>
        /// Returns true if every cell in the row has null values
        /// </summary>
        public bool IsEmpty
        {
            get
            {
                return !_rowRange.Any(x => x.Value != null);
            }
        }

        /// <summary>
        /// An <see cref="ExcelRangeBase"/> representing the row from first to last cell
        /// </summary>
        public ExcelRangeBase RowRange
        {
            get { return _rowRange; }
        }

        /// <summary>
        /// Returns cell value by column name
        /// </summary>
        /// <typeparam name="T">Cell value type</typeparam>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public T GetValue<T>(string columnName)
        {
            var ix = _table.Columns.GetIndexOfColName(columnName);
            return _table.WorkSheet.Cells[_rowIx, _table.Range._fromCol + ix].GetValue<T>();
        }

        /// <summary>
        /// Returns cell value by column name
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public object GetValue(string columnName)
        {
            var ix = _table.Columns.GetIndexOfColName(columnName);
            return _table.WorkSheet.Cells[_rowIx, _table.Range._fromCol + ix].GetValue<object>();
        }

        /// <summary>
        /// Returns cell value by column index
        /// </summary>
        /// <typeparam name="T">Cell value type</typeparam>
        /// <param name="offsetIndex">0-based column index</param>
        /// <returns></returns>
        public T GetValue<T>(int offsetIndex)
        {
            return _table.WorkSheet.Cells[_rowIx, _table.Range._fromCol + offsetIndex].GetValue<T>();
        }

        /// <summary>
        /// Returns cell value by column index
        /// </summary>
        /// <param name="offsetIndex">0-based column index</param>
        /// <returns></returns>
        public object GetValue(int offsetIndex)
        {
            return _table.WorkSheet.Cells[_rowIx, _table.Range._fromCol + offsetIndex].GetValue<object>();
        }

        /// <summary>
        /// Set a cell value by column name
        /// </summary>
        /// <param name="columnName">The table column name</param>
        /// <param name="value">The table cell value</param>
        /// <returns></returns>
        public ExcelTableRow SetValue(string columnName, object value)
        {
            var ix = _table.Columns.GetIndexOfColName(columnName);
            _table.WorkSheet.Cells[_rowIx, _table.Range._fromCol + ix].Value = value;
            return this;
        }

        /// <summary>
        /// Set a cell value by column index
        /// </summary>
        /// <param name="offsetIndex">0-based column index</param>
        /// <param name="value">The table cell value</param>
        /// <returns></returns>
        public ExcelTableRow SetValue(int offsetIndex, object value)
        {
            _table.WorkSheet.Cells[_rowIx, _table.Range._fromCol + offsetIndex].Value = value;
            return this;
        }

        /// <summary>
        /// Clear all cell values in the row's range.
        /// </summary>
        /// <returns></returns>
        public ExcelTableRow Clear()
        {
            _rowRange.Clear();
            return this;
        }
    }
}
