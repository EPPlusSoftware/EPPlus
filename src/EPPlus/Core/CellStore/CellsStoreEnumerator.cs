/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System.Collections.Generic;
using System.Collections;

namespace OfficeOpenXml.Core.CellStore
{
    internal class CellStoreEnumerator<T> : IEnumerable<T>, IEnumerator<T>
    {
        CellStore<T> _cellStore;
        int row, colPos;
        internal int _startRow, _startCol, _endRow, _endCol;
        int minRow, minColPos, maxRow, maxColPos;
        int lastColCount;
        public CellStoreEnumerator(CellStore<T> cellStore) :
            this(cellStore, 0, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns)
        {
        }
        public CellStoreEnumerator(CellStore<T> cellStore, int StartRow, int StartCol, int EndRow, int EndCol)
        {
            _cellStore = cellStore;

            _startRow = StartRow;
            _startCol = StartCol;
            _endRow = EndRow;
            _endCol = EndCol;

            Init();
        }

        internal void Init()
        {
            minRow = _startRow;
            maxRow = _endRow;

            UpdateMinMaxColPos();
            lastColCount = _cellStore.ColumnCount;
            row = minRow;
            colPos = minColPos - 1;
        }

        private void UpdateMinMaxColPos()
        {
            minColPos = _cellStore.GetColumnPosition(_startCol);
            if (minColPos < 0) minColPos = ~minColPos;
            maxColPos = _cellStore.GetColumnPosition(_endCol);
            if (maxColPos < 0) maxColPos = ~maxColPos - 1;
            lastColCount = _cellStore.ColumnCount;  
        }

        internal int Row
        {
            get
            {
                return row;
            }
        }
        internal int Column
        {
            get
            {
                if (colPos<0 || colPos>=_cellStore.ColumnCount)
                {
                    return -1;
                }
                return _cellStore._columnIndex[colPos].Index;
            }
        }
        internal T Value
        {
            get
            {
                lock (_cellStore)
                {
                    return _cellStore.GetValue(row, Column);
                }
            }
            set
            {
                lock (_cellStore)
                {
                    _cellStore.SetValue(row, Column, value);
                }
            }
        }
        internal bool Next()
        {
            if (lastColCount != _cellStore.ColumnCount) UpdateMinMaxColPos();
            return _cellStore.GetNextCell(ref row, ref colPos, minColPos, maxRow, maxColPos);
        }
        internal bool Previous()
        {
            lock (_cellStore)
            {
                if (lastColCount != _cellStore.ColumnCount) UpdateMinMaxColPos();
                return _cellStore.GetPrevCell(ref row, ref colPos, minRow, minColPos, maxColPos);
            }
        }

        public string CellAddress
        {
            get
            {
                return ExcelAddressBase.GetAddress(Row, Column);
            }
        }

        public IEnumerator<T> GetEnumerator()
        {
            Reset();
            return this;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            Reset();
            return this;
        }

        public T Current
        {
            get
            {
                return Value;
            }
        }

        public void Dispose()
        {

        }

        object IEnumerator.Current
        {
            get
            {
                Reset();
                return this;
            }
        }

        public bool MoveNext()
        {
            return Next();
        }

        public void Reset()
        {
            Init();
        }
    }
}