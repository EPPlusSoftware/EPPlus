using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table
{
    /// <summary>
    /// A collection of table rows
    /// </summary>
    public class ExcelTableRowCollection : IEnumerable<ExcelTableRow>
    {
        internal ExcelTableRowCollection(ExcelTable table) 
        {
            _table = table;
        }

        private readonly ExcelTable _table;

        internal event EventHandler<RowsDeletedEventArgs> RowsDeleted;

        private void OnRowsDeleted(int nRows, int position)
        {
            RowsDeleted?.Invoke(this, new RowsDeletedEventArgs(nRows, position));
        }

        /// <summary>
        /// Returns a table row
        /// </summary>
        /// <param name="ix"></param>
        /// <returns></returns>
        public ExcelTableRow this[int ix]
        {
            get
            {
                return new ExcelTableRow(_table, ix);
            }
        }

        /// <summary>
        /// Add a new empty row at the bottom the table.
        /// </summary>
        /// <param name="copyStyles"></param>
        /// <returns></returns>
        public ExcelTableRow AddNewRow(bool copyStyles = true)
        {
            var r = _table.InsertRow(this.Count(), 1, copyStyles);
            return new ExcelTableRow(_table, r.Start.Row - _table.DataRange.Start.Row);
        }

        /// <summary>
        /// Add a number of new empty rows at the bottom the table.
        /// </summary>
        /// <param name="nRows">Number of rows to add</param>
        /// <param name="copyStyles"></param>
        /// <returns></returns>
        public IEnumerable<ExcelTableRow> AddNewRows(int nRows, bool copyStyles = true)
        {
            var rowIx = this.Count();
            _table.InsertRow(rowIx, nRows, copyStyles);
            var result = new List<ExcelTableRow>();
            for (var r = rowIx; r < rowIx + nRows; r++)
            {
                result.Add(new ExcelTableRow(_table, r));
            }
            return result;
        }

        /// <summary>
        /// Inserts a new empty row at the specified position.
        /// </summary>
        /// <param name="position"></param>
        /// <param name="copyStyles"></param>
        /// <returns></returns>
        public ExcelTableRow InsertNewRow(int position, bool copyStyles = true)
        {
            var r = _table.InsertRow(position, 1, copyStyles);
            return new ExcelTableRow(_table, r.Start.Row - _table.DataRange._fromRow);
        }

        /// <summary>
        /// Inserts one or more new empty rows at the specified position.
        /// </summary>
        /// <param name="position"></param>
        /// <param name="nRows">Number of new rows to insert</param>
        /// <param name="copyStyles"></param>
        /// <returns></returns>
        public IEnumerable<ExcelTableRow> InsertNewRows(int position, int nRows = 1, bool copyStyles = true)
        {
            var insertedRange = _table.InsertRow(position, nRows, copyStyles);
            var result = new List<ExcelTableRow>();
            var startRow = insertedRange._fromRow;
            for(var r = startRow; r < startRow + nRows; r++)
            {
                result.Add(new ExcelTableRow(_table, r - _table.DataRange._fromRow));
            }
            return result;
        }

        /// <summary>
        /// Deletes the specified number of rows at the given position in the table
        /// </summary>
        /// <param name="position">0-based position of the deletion</param>
        /// <param name="numberOfRows">Number of rows to delete</param>
        public void DeleteRows(int position, int numberOfRows = 1)
        {
            _table.DeleteRow(position, numberOfRows);
            OnRowsDeleted(numberOfRows, position);
        }

        /// <summary>
        /// Returns an iterator
        /// </summary>
        /// <returns></returns>
        public IEnumerator<ExcelTableRow> GetEnumerator()
        {
            for (int rowNo = _table.DataRange._fromRow; rowNo <= _table.DataRange._toRow; rowNo++)
            {
                yield return new ExcelTableRow(_table, rowNo - _table.DataRange._fromRow);
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <summary>
        /// Clears/deletes all data in the table's rows.
        /// </summary>
        public void Clear()
        {
            _table.DeleteRow(0, this.Count() - 1);
            _table.DataRows[0].Clear();
        }
    }
}
