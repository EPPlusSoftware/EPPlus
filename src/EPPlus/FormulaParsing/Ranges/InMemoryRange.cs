/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/31/2022         EPPlus Software AB           EPPlus 6.1
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.LoadFunctions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Ranges
{
    /// <summary>
    /// EPPlus implementation of a range that keeps its data in memory
    /// </summary>
    public class InMemoryRange : IRangeInfo
    {
        /// <summary>
        /// The constructor
        /// </summary>
        /// <param name="rangeDef">Defines the size of the range</param>
        public InMemoryRange(RangeDefinition rangeDef)
        {
            _cells = new ICellInfo[rangeDef.NumberOfRows, rangeDef.NumberOfCols];
            Size = rangeDef;
            _address = new FormulaRangeAddress() { FromRow = 0, FromCol = 0, ToRow = rangeDef.NumberOfRows - 1, ToCol = rangeDef.NumberOfCols - 1 };
        }
        /// <summary>
        /// The constructor
        /// </summary>
        /// <param name="address">The worksheet address that should be used for this range. Will be used for implicit intersection.</param>
        /// <param name="rangeDef">Defines the size of the range</param>
        public InMemoryRange(FormulaRangeAddress address, RangeDefinition rangeDef)
        {
            if (address?._context != null)
            {
                _ws = address._context.Package.Workbook.Worksheets[address._context.CurrentCell.WorksheetIx];
            }
            _address = address;
            _cells = new ICellInfo[rangeDef.NumberOfRows, rangeDef.NumberOfCols];
            Size = rangeDef;
        }
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="range">A list of values also defining the size of the range</param>
        public InMemoryRange(List<List<object>> range)
        {
            Size = new RangeDefinition(range.Count, (short)range[0].Count);
            _cells = new ICellInfo[Size.NumberOfRows, Size.NumberOfCols];
            for(int c=0;c < Size.NumberOfCols; c++)
            {
                for(int r=0;r< Size.NumberOfRows; r++)
                {
                    _cells[r, c] = new InMemoryCellInfo(range[r][c]);
                }
            }
            _address = new FormulaRangeAddress() { FromRow = 0, FromCol = 0, ToRow = Size.NumberOfRows - 1, ToCol = Size.NumberOfCols - 1 };
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="ri">Another <see cref="IRangeInfo"/> used as clone for this range. The address of the supplied range will not be copied.</param>
        public InMemoryRange(IRangeInfo ri)
        {
            Size = ri.Size;
            _cells = new ICellInfo[Size.NumberOfRows, Size.NumberOfCols];
            for (int c = 0; c < Size.NumberOfCols; c++)
            {
                for (int r = 0; r < Size.NumberOfRows; r++)
                {
                    _cells[r, c] = new InMemoryCellInfo(ri.GetOffset(r, c));
                }
            }
            _address = new FormulaRangeAddress() { FromRow = 0, FromCol = 0, ToRow = Size.NumberOfRows - 1, ToCol = Size.NumberOfCols - 1 };
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="rows">Number of rows in the new range</param>
        /// <param name="cols">Number of columns in the new range</param>
        public InMemoryRange(int rows, short cols)
            : this(new RangeDefinition(rows, cols))
        {
        }

        private readonly FormulaRangeAddress _address;
        private readonly ExcelWorksheet _ws;
        private readonly ICellInfo[,] _cells;
        private int _colIx = -1;
        private int _rowIndex = 0;

        private static InMemoryRange _empty = new InMemoryRange(new RangeDefinition(0, 0));
        
        /// <summary>
        /// An empty range
        /// </summary>
        public static InMemoryRange Empty => _empty;

        /// <summary>
        /// Sets the value for a cell.
        /// </summary>
        /// <param name="row">The row</param>
        /// <param name="col">The column</param>
        /// <param name="val">The value to set</param>
        public void SetValue(int row, int col, object val)
        {
            var c = new InMemoryCellInfo(val);
            _cells[row, col] = c;
        }

        /// <summary>
        /// Sets the <see cref="ICellInfo"/> for a cell directly
        /// </summary>
        /// <param name="row">The row</param>
        /// <param name="col">The column</param>
        /// <param name="cell">The cell</param>
        public void SetCell(int row, int col, ICellInfo cell)
        {
            _cells[row, col] = cell;
        }
        /// <summary>
        /// The in-memory range is never a reference error. Allways false.
        /// </summary>
        public bool IsRef => false;
        /// <summary>
        /// If the range has no cells.
        /// </summary>
        public bool IsEmpty => _cells.Length == 0;
        /// <summary>
        /// If the range is more than one cell.
        /// </summary>
        public bool IsMulti => Size.NumberOfRows * Size.NumberOfCols > 1;
        /// <summary>
        /// If the range is an inmemory range. Allways true.
        /// </summary>
        public bool IsInMemoryRange => true;
        /// <summary>
        /// The size of the range.
        /// </summary>
        public RangeDefinition Size { get; private set; }
        /// <summary>
        /// The address of the inmemory range.
        /// </summary>
        public FormulaRangeAddress Address => _address;
        /// <summary>
        /// The worksheet.
        /// </summary>
        public ExcelWorksheet Worksheet => _ws;
        /// <summary>
        /// The address of the range
        /// </summary>
        public FormulaRangeAddress Dimension
        {
            get
            {        
                return _address;
            }
        }
        /// <summary>
        /// Current
        /// </summary>
        public ICellInfo Current
        {
            get
            {
                return _cells[_rowIndex, _colIx] ?? new InMemoryCellInfo(null);
            }
        }

        object IEnumerator.Current
        {
            get
            {
                return _cells[_rowIndex, _colIx] ?? new InMemoryCellInfo(null);
            }
        }
        /// <summary>
        /// Dispose
        /// </summary>
        public void Dispose()
        {

        }
        /// <summary>
        /// Get enumerator
        /// </summary>
        /// <returns></returns>
        public IEnumerator<ICellInfo> GetEnumerator()
        {
            _colIx = -1;
            _rowIndex = 0;
            return this;
        }

        /// <summary>
        /// Get the number of cells in the range
        /// </summary>
        /// <returns>The number of cells in range.</returns>
        public int GetNCells()
        {
            return Size.NumberOfRows * Size.NumberOfCols;
        }
        /// <summary>
        /// Returns the value with the offset from the top-left cell.
        /// </summary>
        /// <param name="rowOffset">The row offset from the top-left cell.</param>
        /// <param name="colOffset">The column offset from the top-left cell.</param>
        /// <returns>The value of the cell</returns>
        public object GetOffset(int rowOffset, int colOffset)
        {
            var c = _cells[rowOffset, colOffset];
            if (c == null)
            {
                return null;
            }
            return c.Value;
        }

        /// <summary>
        /// Returns the value with the offset from the top-left cell.
        /// </summary>
        /// <param name="rowOffsetStart">The starting row offset from the top-left cell.</param>
        /// <param name="colOffsetEnd">The starting column offset from the top-left cell.</param>
        /// <param name="rowOffsetEnd">The ending row offset from the top-left cell.</param>
        /// <param name="colOffsetStart">The ending column offset from the top-left cell</param>
        /// <returns>The value of the cell</returns>
        public IRangeInfo GetOffset(int rowOffsetStart, int colOffsetStart, int rowOffsetEnd, int colOffsetEnd)
        {
            var nRows = Math.Abs(rowOffsetEnd - rowOffsetStart);
            var nCols = (short)Math.Abs(colOffsetEnd- colOffsetStart);
            nRows++;
            nCols++;
            var rangeDef = new RangeDefinition(nRows, nCols);
            var result = new InMemoryRange(rangeDef);
            var rowIx = 0;
            for(var row = rowOffsetStart; row <= rowOffsetEnd; row++)
            {
                var colIx = 0;
                for (var col = colOffsetStart; col <= colOffsetEnd; col++)
                {
                    result.SetValue(rowIx, colIx++, _cells[row, col].Value);
                }
                rowIx++;
            }
            return result;
        }
        /// <summary>
        /// If the cell's row is hidden.
        /// </summary>
        /// <param name="rowOffset">Row offset from the top-left cell</param>
        /// <param name="colOffset">Column offset from the top-left cell</param>
        /// <returns></returns>
        public bool IsHidden(int rowOffset, int colOffset)
        {
            return false;
        }
        /// <summary>
        /// Gets the value of a cell.
        /// </summary>
        /// <param name="row">The row</param>
        /// <param name="col">The column</param>
        /// <returns></returns>
        public object GetValue(int row, int col)
        {
            if (_address == null)
            { 
                var c = _cells[row, col];
                if (c == null) return null;
                return c.Value;


            }
            else
            {
                var c = _cells[row-_address.FromRow, col-Address.FromCol];
                if (c == null) return null;
                return c.Value;
            }
        }
        /// <summary>
        /// Get cell
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public ICellInfo GetCell(int row, int col)
        {
            var c = _cells[row, col];
            if (c == null) return null;
            return c;
        }
        /// <summary>
        /// Move next
        /// </summary>
        /// <returns></returns>
        public bool MoveNext()
        {
            if (_colIx < Size.NumberOfCols - 1)
            {
                _colIx++;
                return true;
            }
            _colIx = 0;
            _rowIndex++;
            if (_rowIndex >= Size.NumberOfRows) return false;
            return true;
        }
        /// <summary>
        /// Reset
        /// </summary>
        public void Reset()
        {
            _colIx = 0;
            _rowIndex = 0;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            _colIx = -1;
            _rowIndex = 0;
            return this;
        }

        internal static InMemoryRange CloneRange(IRangeInfo ri)
        {
            var ret = new InMemoryRange(ri.Size);
            for(int r=0;r < ri.Size.NumberOfRows;r++)
            {
                for (int c = 0; c < ri.Size.NumberOfCols;c++)
                {
                    ret.SetValue(r,c,ri.GetOffset(r,c));
                }
            }

            return ret;
        }

        internal static InMemoryRange GetFromArray(params object[] values)
        {
            var rows = values.GetUpperBound(0) + 1;
            var ir = new InMemoryRange(rows, 1);
            for(int r=0;r < rows;r++)
            {
                ir.SetValue(r, 0, values[r]);
            }
            return ir;
        }
    }
}
