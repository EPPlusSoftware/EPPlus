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
    internal class InMemoryRange : IRangeInfo
    {
        public InMemoryRange(RangeDefinition rangeDef)
        {
            _cells = new ICellInfo[rangeDef.NumberOfRows, rangeDef.NumberOfCols];
            _size = rangeDef;
            _address = new FormulaRangeAddress() { FromRow = 0, FromCol = 0, ToRow = rangeDef.NumberOfRows - 1, ToCol = rangeDef.NumberOfCols - 1 };
        }
        public InMemoryRange(FormulaRangeAddress address, RangeDefinition rangeDef, ParsingContext ctx)
        {
            _ws = ctx.Package.Workbook.Worksheets[ctx.CurrentCell.WorksheetIx];
            _address = address;
            _cells = new ICellInfo[rangeDef.NumberOfRows, rangeDef.NumberOfCols];
            _size = rangeDef;
        }
        public InMemoryRange(List<List<object>> range)
        {
            _size = new RangeDefinition(range.Count, (short)range[0].Count);
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

        public InMemoryRange(IRangeInfo ri)
        {
            var size = ri.Size;
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

        public InMemoryRange(int rows, short cols)
            : this(new RangeDefinition(rows, cols))
        {}

        private readonly FormulaRangeAddress _address;
        private readonly RangeDefinition _size;
        private readonly ExcelWorksheet _ws;
        private readonly ICellInfo[,] _cells;
        private int _colIx = -1;
        private int _rowIndex = 0;
        //private readonly short _nCols;
        //private readonly int _nRows;

        private static InMemoryRange _empty = new InMemoryRange(new RangeDefinition(0, 0));
        
        /// <summary>
        /// An empty range
        /// </summary>
        public static InMemoryRange Empty => _empty;

        public void SetValue(int row, int col, object val)
        {
            var c = new InMemoryCellInfo(val);
            _cells[row, col] = c;
        }

        public void SetCell(int row, int col, ICellInfo cell)
        {
            _cells[row, col] = cell;
        }
        public bool IsRef => false;
        public bool IsEmpty => _cells.Length == 0;

        public bool IsMulti => Size.NumberOfRows * Size.NumberOfCols > 1;

        public bool IsInMemoryRange => true;

        public RangeDefinition Size => _size;

        public FormulaRangeAddress Address => _address;

        public ExcelWorksheet Worksheet => _ws;
        public FormulaRangeAddress Dimension
        {
            get
            {        
                return _address;
            }
        }
        public ICellInfo Current
        {
            get
            {
                return _cells[_rowIndex, _colIx];
            }
        }

        object IEnumerator.Current
        {
            get
            {
                return _cells[_rowIndex, _colIx];
            }
        }

        public void Dispose()
        {

        }

        public IEnumerator<ICellInfo> GetEnumerator()
        {
            _colIx = -1;
            _rowIndex = 0;
            return this;
        }

        public int GetNCells()
        {
            return _size.NumberOfRows * _size.NumberOfCols;
        }

        public object GetOffset(int rowOffset, int colOffset)
        {
            var c = _cells[rowOffset, colOffset];
            if (c == null)
            {
                return null;
            }
            return c.Value;
        }

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
        public bool IsHidden(int rowOffset, int colOffset)
        {
            return false;
        }
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

        public ICellInfo GetCell(int row, int col)
        {
            var c = _cells[row, col];
            if (c == null) return null;
            return c;
        }

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
    }
}
