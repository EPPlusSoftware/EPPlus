using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table
{
    /// <summary>
    /// Represents a row in an <see cref="ExcelTable"/>
    /// </summary>
    public class ExcelTableRow
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="table">The <see cref="ExcelTable"/> that the row belongs to</param>
        /// <param name="rowIx">0 based index of the table row</param>
        internal ExcelTableRow(ExcelTable table, int rowIx)
        {
            _table = table;
            _rowIx = rowIx;
            _table.DataRows.RowsDeleted += DataRows_RowsDeleted;
            SetRowRange();
        }

        private void SetRowRange()
        {
            var startCol = _table.DataRange._fromCol;
            var endCol = _table.DataRange._toCol;
            var startRow = _table.DataRange._fromRow;
            _rowRange = _table.WorkSheet.Cells[startRow + _rowIx, startCol, startRow + _rowIx, endCol];
        }

        private void DataRows_RowsDeleted(object sender, RowsDeletedEventArgs e)
        {
            if(e.Position <= _rowIx && _rowIx <= e.Position + e.NumberOfDeletedRows - 1)
            {
                _isDeleted = true;
                _table.DataRows.RowsDeleted -= DataRows_RowsDeleted;
            }
            else if(_rowIx > e.Position + e.NumberOfDeletedRows - 1)
            {
                _rowIx -= e.NumberOfDeletedRows;
                SetRowRange();
            }
        }

        private readonly ExcelTable _table;
        private int _rowIx;
        private ExcelRangeBase _rowRange;
        private bool _isDeleted;

        private void CheckDeleted()
        {
            if(_isDeleted)
            {
                throw new InvalidOperationException("Cannot call methods/properties on deleted table row");
            }
        }

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
                return _table.WorkSheet.Row(_rowRange.Start.Row).Hidden;
            }
        }

        /// <summary>
        /// Indicates if this row has been deleted.
        /// </summary>
        public bool IsDeleted => _isDeleted;

        internal int RowIx => _rowIx;


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
        /// Returns formula by column name.
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public string GetFormula(string columnName)
        {
            CheckDeleted();
            var ix = _table.Columns.GetIndexOfColName(columnName);
            var startRow = _table.DataRange._fromRow;
            return _table.WorkSheet.Cells[startRow + _rowIx, _table.Range._fromCol + ix].Formula;
        }

        /// <summary>
        /// Returns formula by 0-based column index
        /// </summary>
        /// <param name="offsetIndex"></param>
        /// <returns></returns>
        public string GetFormula(int offsetIndex)
        {
            CheckDeleted();
            var startRow = _table.DataRange._fromRow;
            return _table.WorkSheet.Cells[startRow + _rowIx, _table.Range._fromCol + offsetIndex].Formula;
        }

        /// <summary>
        /// Returns cell value by column name
        /// </summary>
        /// <typeparam name="T">Cell value type</typeparam>
        /// <param name="columnName"></param>
        /// <exception cref="InvalidOperationException">If the row has previously been deleted.</exception>
        /// <returns></returns>
        public T GetValue<T>(string columnName)
        {
            CheckDeleted();
            var ix = _table.Columns.GetIndexOfColName(columnName);
            var startRow = _table.DataRange._fromRow;
            return _table.WorkSheet.Cells[startRow + _rowIx, _table.Range._fromCol + ix].GetValue<T>();
        }

        /// <summary>
        /// Returns cell value by column name
        /// </summary>
        /// <param name="columnName"></param>
        /// <exception cref="InvalidOperationException">If the row has previously been deleted.</exception>
        /// <returns></returns>
        public object GetValue(string columnName)
        {
            CheckDeleted();
            var ix = _table.Columns.GetIndexOfColName(columnName);
            var startRow = _table.DataRange._fromRow;
            return _table.WorkSheet.Cells[startRow + _rowIx, _table.Range._fromCol + ix].GetValue<object>();
        }

        /// <summary>
        /// Returns cell value by column index
        /// </summary>
        /// <typeparam name="T">Cell value type</typeparam>
        /// <param name="offsetIndex">0-based column index</param>
        /// <exception cref="InvalidOperationException">If the row has previously been deleted.</exception>
        /// <returns></returns>
        public T GetValue<T>(int offsetIndex)
        {
            CheckDeleted();
            var startRow = _table.DataRange._fromRow;
            return _table.WorkSheet.Cells[startRow + _rowIx, _table.Range._fromCol + offsetIndex].GetValue<T>();
        }

        /// <summary>
        /// Returns cell value by column index
        /// </summary>
        /// <param name="offsetIndex">0-based column index</param>
        /// <exception cref="InvalidOperationException">If the row has previously been deleted.</exception>
        /// <returns></returns>
        public object GetValue(int offsetIndex)
        {
            CheckDeleted();
            var startRow = _table.DataRange._fromRow;
            return _table.WorkSheet.Cells[startRow + _rowIx, _table.Range._fromCol + offsetIndex].GetValue<object>();
        }

        /// <summary>
        /// Set a cell value by column name
        /// </summary>
        /// <param name="columnName">The table column name</param>
        /// <param name="value">The table cell value</param>
        /// <exception cref="InvalidOperationException">If the row has previously been deleted.</exception>
        /// <returns></returns>
        public ExcelTableRow SetValue(string columnName, object value)
        {
            CheckDeleted();
            var ix = _table.Columns.GetIndexOfColName(columnName);
            var startRow = _table.DataRange._fromRow;
            _table.WorkSheet.Cells[startRow + _rowIx, _table.Range._fromCol + ix].Value = value;
            return this;
        }

        /// <summary>
        /// Set a cell value by column index
        /// </summary>
        /// <param name="offsetIndex">0-based column index</param>
        /// <param name="value">The table cell value</param>
        /// <exception cref="InvalidOperationException">If the row has previously been deleted.</exception>
        /// <returns></returns>
        public ExcelTableRow SetValue(int offsetIndex, object value)
        {
            CheckDeleted();
            var startRow = _table.DataRange._fromRow;
            _table.WorkSheet.Cells[startRow + _rowIx, _table.Range._fromCol + offsetIndex].Value = value;
            return this;
        }

        /// <summary>
        /// Set all the cell values of the table row by providing an array of <see cref="object"/>
        /// </summary>
        /// <param name="values"></param>
        /// <exception cref="ArgumentNullException">Will be thrown if <paramref name="values"/> is null</exception>
        /// <exception cref="ArgumentException">Will be thrown if <paramref name="values"/> is an empty array</exception>
        /// <exception cref="ArgumentOutOfRangeException">Will be thrown if number of items in <paramref name="values"/> exceeds number of columns in the table.</exception>
        /// <exception cref="InvalidOperationException">If the row has previously been deleted.</exception>
        public void SetValues(params object[] values)
        {
            CheckDeleted();
            if (values == null) throw new ArgumentNullException(nameof(values));
            if(values.Length == 0) throw new ArgumentException("values cannot be an empty array", nameof(values));
            var vals = values;
            if (values[0].GetType().IsArray)
            {
                vals = ((IEnumerable)values[0]).Cast<object>().ToArray();
            }
            var row = _rowRange._fromRow;
            var startCol = _rowRange._fromCol;
            var endCol = _rowRange._toCol;
            if (vals.Length > endCol - startCol + 1) throw new ArgumentOutOfRangeException(nameof(values), "Number of values exceeds number of columns in the table");
            for(var col = startCol; col <= endCol; col++)
            {
                if (col - startCol > vals.Length) break;
                _rowRange.Worksheet.Cells[row, col].Value = vals[col - startCol];
            }
        }

        /// <summary>
        /// Removes this row from the table
        /// </summary>
        public void Delete()
        {
            if(!_isDeleted)
            {
                _table.DataRows.DeleteRows(_rowIx, 1);
            }
        }

        /// <summary>
        /// Clear all cell values in the row's range.
        /// </summary>
        /// <returns></returns>
        public ExcelTableRow Clear()
        {
            CheckDeleted();
            _rowRange.Clear();
            return this;
        }
    }
}
