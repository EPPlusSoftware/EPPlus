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
            _nRows = rangeDef.NumberOfRows;
            _nCols = rangeDef.NumberOfCols;
            _cells = new ICellInfo[_nCols, _nRows];
            _size = rangeDef;
        }
        public InMemoryRange(FormulaRangeAddress address, RangeDefinition rangeDef, ParsingContext ctx)
        {
            _ws = ctx.Package.Workbook.Worksheets[ctx.CurrentCell.WorksheetIx];
            _address = address;
            _nRows = rangeDef.NumberOfRows;
            _nCols = rangeDef.NumberOfCols;
            _cells = new ICellInfo[_nCols, _nRows];
            _size = rangeDef;
        }

        private readonly FormulaRangeAddress _address;
        private readonly RangeDefinition _size;
        private readonly ExcelWorksheet _ws;
        private readonly ICellInfo[,] _cells;
        private int _colIx = -1;
        private int _rowIndex = 0;
        private readonly short _nCols;
        private readonly int _nRows;

        private static InMemoryRange _empty = new InMemoryRange(new RangeDefinition(0, 0));
        
        /// <summary>
        /// An empty range
        /// </summary>
        public static InMemoryRange Empty => _empty;

        public void SetValue(int row, int col, object val)
        {
            var c = new InMemoryCellInfo(val);
            _cells[col, row] = c;
        }

        public void SetCell(int row, int col, ICellInfo cell)
        {
            _cells[col, row] = cell;
        }

        public bool IsEmpty => _cells.Length == 0;

        public bool IsMulti => _cells.Length > 0 && _cells.GetUpperBound(1) > 1;

        public bool IsInMemoryRange => true;

        public RangeDefinition Size => _size;

        public FormulaRangeAddress Address => _address;

        public ExcelWorksheet Worksheet => _ws;

        public ICellInfo Current
        {
            get
            {
                return _cells[_colIx, _rowIndex];
            }
        }

        object IEnumerator.Current
        {
            get
            {
                return _cells[_colIx, _rowIndex];
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
            var c = _cells[colOffset, rowOffset];
            if (c == null)
            {
                return null;
            }
            return c.Value;
        }

        public object GetValue(int row, int col)
        {
            var c = _cells[col, row];
            if (c == null) return null;
            return c.Value;
        }

        public ICellInfo GetCell(int row, int col)
        {
            var c = _cells[col, row];
            if (c == null) return null;
            return c;
        }

        public bool MoveNext()
        {
            if (_colIx < _nCols - 1)
            {
                _colIx++;
                return true;
            }
            _colIx = 0;
            _rowIndex++;
            if (_rowIndex >= _nRows) return false;
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
    }
}
