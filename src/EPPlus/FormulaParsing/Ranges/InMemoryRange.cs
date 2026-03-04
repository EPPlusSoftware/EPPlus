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
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Ranges
{
    /// <summary>
    /// EPPlus implementation of a range that keeps its data in memory
    /// </summary>
    [DebuggerDisplay("{Size}")]
    public class InMemoryRange : IRangeInfo
    {

        /// <summary>
        /// The constructor
        /// </summary>
        /// <param name="rangeDef">Defines the size of the range</param>
        public InMemoryRange(RangeDefinition rangeDef)
        {
            _cells = new ICellInfo[rangeDef.NumberOfRows, rangeDef.NumberOfCols];
            _physicalRows = rangeDef.NumberOfRows;
            Size = rangeDef;
            _address = new FormulaRangeAddress() { FromRow = 0, FromCol = 0, ToRow = rangeDef.NumberOfRows - 1, ToCol = rangeDef.NumberOfCols - 1 };
            Debug.Assert(_physicalRows >= 0);
            Debug.Assert(_physicalRows <= Size.NumberOfRows);
        }
        /// <summary>
        /// The constructor
        /// </summary>
        /// <param name="address">The worksheet address that should be used for this range.
        /// Will be used for implicit intersection.</param>
        /// <param name="rangeDef">Defines the size of the range</param>
        public InMemoryRange(FormulaRangeAddress address, RangeDefinition rangeDef)
        {
            if (address != null && address._context != null)
            {
                _ws = address._context.Package.Workbook.GetWorksheetByIndexInList(address._context.CurrentCell.WorksheetIx);
            }
            _address = address;
            _cells = new ICellInfo[rangeDef.NumberOfRows, rangeDef.NumberOfCols];
            _physicalRows = rangeDef.NumberOfRows;
            Size = rangeDef;
            Debug.Assert(_physicalRows >= 0);
            Debug.Assert(_physicalRows <= Size.NumberOfRows);
        }

        /// <summary>
        /// Constructor for a virtual range with a small physical backing array but large logical size.
        /// Rows beyond <paramref name="physicalDef"/>.NumberOfRows are virtual and return null.
        /// </summary>
        /// <param name="address">The worksheet address for implicit intersection.</param>
        /// <param name="physicalDef">Defines the physical (backed) size of the range.</param>
        /// <param name="logicalRows">The full logical row count (e.g. 1048576 for a full column).</param>
        internal InMemoryRange(FormulaRangeAddress address, RangeDefinition physicalDef, int logicalRows)
        {
            if (address != null && address._context != null)
            {
                _ws = address._context.Package.Workbook.GetWorksheetByIndexInList(address._context.CurrentCell.WorksheetIx);
            }
            _address = address;
            _physicalRows = physicalDef.NumberOfRows;
            _cells = new ICellInfo[physicalDef.NumberOfRows, physicalDef.NumberOfCols];
            Size = new RangeDefinition(logicalRows, physicalDef.NumberOfCols);
            Debug.Assert(_physicalRows >= 0);
            Debug.Assert(_physicalRows <= Size.NumberOfRows);
        }

        /// <summary>
        /// Constructor for a virtual range without an address.
        /// Rows beyond <paramref name="physicalDef"/>.NumberOfRows are virtual and return null.
        /// </summary>
        /// <param name="physicalDef">Defines the physical (backed) size of the range.</param>
        /// <param name="logicalRows">The full logical row count.</param>
        internal InMemoryRange(RangeDefinition physicalDef, int logicalRows)
        {
            _physicalRows = physicalDef.NumberOfRows;
            _cells = new ICellInfo[physicalDef.NumberOfRows, physicalDef.NumberOfCols];
            Size = new RangeDefinition(logicalRows, physicalDef.NumberOfCols);
            _address = new FormulaRangeAddress()
            {
                FromRow = 0,
                FromCol = 0,
                ToRow = logicalRows - 1,
                ToCol = physicalDef.NumberOfCols - 1
            };
            Debug.Assert(_physicalRows >= 0);
            Debug.Assert(_physicalRows <= Size.NumberOfRows);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="range">A list of values also defining the size of the range</param>
        public InMemoryRange(List<List<object>> range)
        {
            Size = new RangeDefinition(range.Count, (short)range[0].Count);
            _physicalRows = Size.NumberOfRows;
            _cells = new ICellInfo[Size.NumberOfRows, Size.NumberOfCols];
            for (int c = 0; c < Size.NumberOfCols; c++)
            {
                for (int r = 0; r < Size.NumberOfRows; r++)
                {
                    _cells[r, c] = new InMemoryCellInfo(range[r][c]);
                }
            }
            _address = new FormulaRangeAddress() { FromRow = 0, FromCol = 0, ToRow = Size.NumberOfRows - 1, ToCol = Size.NumberOfCols - 1 };
            Debug.Assert(_physicalRows >= 0);
            Debug.Assert(_physicalRows <= Size.NumberOfRows);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="ri">Another <see cref="IRangeInfo"/> used as clone for this range.
        /// The address of the supplied range will not be copied.</param>
        public InMemoryRange(IRangeInfo ri)
        {
            Size = ri.Size;
            _physicalRows = Size.NumberOfRows;
            _cells = new ICellInfo[Size.NumberOfRows, Size.NumberOfCols];
            for (int c = 0; c < Size.NumberOfCols; c++)
            {
                for (int r = 0; r < Size.NumberOfRows; r++)
                {
                    _cells[r, c] = new InMemoryCellInfo(ri.GetOffset(r, c));
                }
            }
            _address = new FormulaRangeAddress() { FromRow = 0, FromCol = 0, ToRow = Size.NumberOfRows - 1, ToCol = Size.NumberOfCols - 1 };
            Debug.Assert(_physicalRows >= 0);
            Debug.Assert(_physicalRows <= Size.NumberOfRows);
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
        private readonly int _physicalRows;
        private object _virtualDefaultValue;  // null means "no default" (backward compat)

        private static InMemoryRange _empty = new InMemoryRange(new RangeDefinition(0, 0));

        /// <summary>
        /// An empty range
        /// </summary>
        public static InMemoryRange Empty => _empty;

        /// <summary>
        /// Number of rows backed by the physical cell array.
        /// For non-virtual ranges this equals Size.NumberOfRows.
        /// </summary>
        internal int PhysicalRows
        {
            get { return _physicalRows; }
        }


        /// <summary>
        /// The value returned for rows beyond PhysicalRows.
        /// When null (default), virtual rows return null.
        /// When set, virtual rows return this value instead.
        /// </summary>
        internal object VirtualDefaultValue
        {
            get { return _virtualDefaultValue; }
            set { _virtualDefaultValue = value; }
        }

        /// <summary>
        /// True if the range has virtual (unstored) rows beyond PhysicalRows.
        /// </summary>
        internal bool HasVirtualRows
        {
            get { return Size.NumberOfRows > _physicalRows; }
        }

        /// <summary>
        /// Sets the value for a cell.
        /// </summary>
        /// <param name="row">The row</param>
        /// <param name="col">The column</param>
        /// <param name="val">The value to set</param>
        public void SetValue(int row, int col, object val)
        {
            if (row >= _physicalRows) return;
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
            if (row >= _physicalRows) return;
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
        /// The addresses for the range, if more than one.
        /// </summary>
        public FormulaRangeAddress[] Addresses => new FormulaRangeAddress[] { _address };

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
            if (rowOffset >= _physicalRows)
                return _virtualDefaultValue;  // was: return null;
            var c = _cells[rowOffset, colOffset];
            return c == null ? null : c.Value;
        }

        /// <summary>
        /// Returns the value with the offset from the top-left cell.
        /// </summary>
        /// <param name="rowOffsetStart">The starting row offset from the top-left cell.</param>
        /// <param name="colOffsetEnd">The starting column offset from the top-left cell.</param>
        /// <param name="rowOffsetEnd">The ending row offset from the top-left cell.</param>
        /// <param name="colOffsetStart">The ending column offset from the top-left cell</param>
        /// <returns>The value of the cell</returns>
        public IRangeInfo GetOffset(int rowOffsetStart, int colOffsetStart,
                             int rowOffsetEnd, int colOffsetEnd)
        {
            var logicalRows = rowOffsetEnd - rowOffsetStart + 1;
            var nCols = (short)(colOffsetEnd - colOffsetStart + 1);

            if (rowOffsetStart >= _physicalRows)
            {
                var emptyRange = new InMemoryRange(new RangeDefinition(0, nCols), logicalRows);
                emptyRange._virtualDefaultValue = _virtualDefaultValue;  // propagate
                return emptyRange;
            }

            var physicalEnd = Math.Min(rowOffsetEnd, _physicalRows - 1);
            var physicalRows = physicalEnd - rowOffsetStart + 1;

            InMemoryRange result;
            if (physicalRows < logicalRows)
            {
                result = new InMemoryRange(new RangeDefinition(physicalRows, nCols), logicalRows);
                result._virtualDefaultValue = _virtualDefaultValue;  // propagate
            }
            else
            {
                result = new InMemoryRange(new RangeDefinition(physicalRows, nCols));
            }

            for (var row = rowOffsetStart; row <= physicalEnd; row++)
            {
                var colIx = 0;
                for (var col = colOffsetStart; col <= colOffsetEnd; col++)
                {
                    var cell = _cells[row, col];
                    var val = cell != null ? cell.Value : null;
                    result.SetValue(row - rowOffsetStart, colIx, val);
                    colIx++;
                }
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
            int r = _address == null ? row : row - _address.FromRow;
            if (r >= _physicalRows) return _virtualDefaultValue;
            int c = _address == null ? col : col - _address.FromCol;
            var cell = _cells[r, c];
            if (cell == null) return null;
            return cell.Value;
        }
        /// <summary>
        /// Get cell
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public ICellInfo GetCell(int row, int col)
        {
            if (row >= _physicalRows)
            {
                // Return a virtual cell with the default value if set
                if (_virtualDefaultValue != null)
                    return new InMemoryCellInfo(_virtualDefaultValue);
                return null;
            }
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
            if (_rowIndex >= _physicalRows) return false;
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
            var isVirtual = ri is InMemoryRange && ((InMemoryRange)ri).HasVirtualRows;
            int physRows = isVirtual ? ((InMemoryRange)ri).PhysicalRows : ri.Size.NumberOfRows;

            var ret = isVirtual
                ? new InMemoryRange(new RangeDefinition(physRows, ri.Size.NumberOfCols),
                                    ri.Size.NumberOfRows)
                : new InMemoryRange(ri.Size);

            if (isVirtual)
            {
                ret._virtualDefaultValue = ((InMemoryRange)ri)._virtualDefaultValue;
            }

            for (int r = 0; r < physRows; r++)
            {
                for (int c = 0; c < ri.Size.NumberOfCols; c++)
                {
                    ret.SetValue(r, c, ri.GetOffset(r, c));
                }
            }
            return ret;
        }

        internal static InMemoryRange GetFromArray(params object[] values)
        {
            var rows = values.GetUpperBound(0) + 1;
            var ir = new InMemoryRange(rows, 1);
            for (int r = 0; r < rows; r++)
            {
                ir.SetValue(r, 0, values[r]);
            }
            return ir;
        }

        /// <summary>
        /// Get the address adjusted inside the dimension of the worksheet.
        /// Not applicable on InMemoryRange's, as no addresses us used.
        /// </summary>
        /// <param name="index">Not applicable on InMemoryRange's.</param>
        /// <returns>The address.</returns>
        public FormulaRangeAddress GetAddressDimensionAdjusted(int index)
        {
            if (index > 0) return null;

            if (HasVirtualRows)
            {
                // Return address clamped to physical rows
                return new FormulaRangeAddress()
                {
                    FromRow = _address.FromRow,
                    FromCol = _address.FromCol,
                    ToRow = _address.FromRow + _physicalRows - 1,
                    ToCol = _address.ToCol
                };
            }
            return _address;
        }

        internal IEnumerable<Token> SerializeToTokens()
        {
            var sb = new StringBuilder();
            sb.Append("{");
            for (var row = 0; row < _physicalRows; row++)
            {
                for (var col = 0; col < Size.NumberOfCols; col++)
                {
                    var val = GetOffset(row, col);
                    sb.Append(val.ToString());
                    if (col < Size.NumberOfCols - 1)
                    {
                        sb.Append(",");
                    }
                }
                if (row < _physicalRows - 1)
                {
                    sb.Append(";");
                }
            }
            sb.Append("}");
            return SourceCodeTokenizer.Default.Tokenize(sb.ToString());
        }
        /// <summary>
        /// Gets the <see cref="IRangeInfo" /> for the range for the first and last value in the range (top-left to bottom-right)
        /// </summary>
        /// <returns>The range</returns>
        public IRangeInfo GetRangeInfoByValue()
        {
            return this;
        }
    }
}