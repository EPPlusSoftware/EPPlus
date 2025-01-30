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

        /// <summary>
        /// Returns a table row
        /// </summary>
        /// <param name="ix"></param>
        /// <returns></returns>
        public ExcelTableRow this[int ix]
        {
            get
            {
                var rowNo = _table.DataRange._fromRow + ix;
                return new ExcelTableRow(_table, rowNo);
            }
        }

        /// <summary>
        /// Add a new empty row at the bottom the table.
        /// </summary>
        /// <param name="copyStyles"></param>
        /// <returns></returns>
        public ExcelTableRow AddNewRow(bool copyStyles = true)
        {
            _table.InsertRow(this.Count(), 1, copyStyles);
            return new ExcelTableRow(_table, _table.DataRange._toRow);
        }

        /// <summary>
        /// Inserts a new empty row at the specified position.
        /// </summary>
        /// <param name="position"></param>
        /// <param name="copyStyles"></param>
        /// <returns></returns>
        public ExcelTableRow InsertNewRow(int position, bool copyStyles = true)
        {
            _table.InsertRow(position, 1, copyStyles);
            return new ExcelTableRow(_table, _table.DataRange._fromRow + position);
        }

        /// <summary>
        /// Deletes the specified number of rows at the given position in the table
        /// </summary>
        /// <param name="position">0-based position of the deletion</param>
        /// <param name="numberOfRows">Number of rows to delete</param>
        public void DeleteRows(int position, int numberOfRows = 1)
        {
            _table.DeleteRow(position, numberOfRows);
        }

        /// <summary>
        /// Returns an iterator
        /// </summary>
        /// <returns></returns>
        public IEnumerator<ExcelTableRow> GetEnumerator()
        {
            for (int rowNo = _table.DataRange._fromRow; rowNo <= _table.DataRange._toRow; rowNo++)
            {
                yield return new ExcelTableRow(_table, rowNo);
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
