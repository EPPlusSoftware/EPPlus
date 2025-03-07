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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.Core.CellStore
{
    internal class CellStoreEnumerator<T> : IEnumerable<T>, IEnumerator<T>
    {
        CellStore<T> _cellStore;
        int _row, _colPos;
        internal int _startRow, _startCol, _endRow, _endCol;
        List<SimpleAddress> _ranges=null;
        int _rangePos = 0;
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
        public CellStoreEnumerator(CellStore<T> cellStore, ExcelAddressBase address)
        {
            _cellStore = cellStore;

            _startRow = address._fromRow;
            _startCol = address._fromCol;
            _endRow = address._toRow;
            _endCol = address._toCol;
            if (address.Addresses != null && address.Addresses.Count > 1)
            {
                _ranges = new List<SimpleAddress>();
                for (int i=0;i < address.Addresses.Count; i++)
                {
                    var a = address._addresses[i];
                    _ranges.Add(new SimpleAddress(a._fromRow, a._fromCol, a._toRow, a._toCol));
                }
            }

            Init();
        }
        public CellStoreEnumerator(CellStore<T> cellStore, FormulaRangeAddress[] addresses)
        {
            _cellStore = cellStore;

            var address = addresses[0];
            _startRow = address.FromRow;
            _startCol = address.FromCol;
            _endRow = address.ToRow;
            _endCol = address.ToCol;

            if (addresses.Length > 1)
            {
                _ranges = new List<SimpleAddress>();
                for (int i = 0; i < addresses.Length; i++)
                {
                    var a = addresses[i];
                    if (a != null)
                    {
                        _ranges.Add(new SimpleAddress(a.FromRow, a.FromCol, a.ToRow, a.ToCol));
                    }
                }
            }
            Init();
        }

        internal void Init()
        {
            _rangePos = 0;
            InitNewRange();
        }

        private void InitNewRange()
        {
            minRow = _startRow;
            maxRow = _endRow;
            UpdateMinMaxColPos();
            lastColCount = _cellStore.ColumnCount;
            _row = minRow;
            _colPos = minColPos - 1;
        }

        private void UpdateMinMaxColPos()
        {
            minColPos = _cellStore.GetColumnPosition(_startCol);
            if (minColPos < 0) minColPos = ~minColPos;
            maxColPos = _cellStore.GetColumnPosition(_endCol);
            if (maxColPos < 0) maxColPos = ~maxColPos - 1;
            lastColCount = _cellStore.ColumnCount;  
        }
        internal int RangePos
        {
            get { return _rangePos; }
        }
        internal List<SimpleAddress> Ranges
        {
            get
            {
                return _ranges;
            }
        }
        internal int Row
        {
            get
            {
                return _row;
            }
        }
        internal int Column
        {
            get
            {
                if (_colPos < 0 || _colPos >= _cellStore.ColumnCount)
                {
                    return -1;
                }
                if(_colPos < minColPos)
                {
                    return _startCol;
                }
                else
                {
                    return _cellStore._columnIndex[_colPos].Index;
                }
            }
        }
        internal T Value
        {
            get
            {
                lock (_cellStore)
                {
                    return _cellStore.GetValue(_row, Column);
                }
            }
            set
            {
                lock (_cellStore)
                {
                    _cellStore.SetValue(_row, Column, value);
                }
            }
        }
        internal bool Next()
        {
            if (lastColCount != _cellStore.ColumnCount) UpdateMinMaxColPos();
            var ret = _cellStore.GetNextCell(ref _row, ref _colPos, minColPos, maxRow, maxColPos);
            if (ret == false && _ranges!=null && _ranges.Count > ++_rangePos)
            {
                var a = _ranges[_rangePos];

                _startRow = a.FromRow;
                _startCol = a.FromCol;
                _endRow = a.ToRow;
                _endCol = a.ToCol;
                
                InitNewRange();                
                return Next();
            }
            return ret;
        }
        internal bool Previous()
        {
            lock (_cellStore)
            {
                if (lastColCount != _cellStore.ColumnCount) UpdateMinMaxColPos();
                var ret = _cellStore.GetPrevCell(ref _row, ref _colPos, minRow, minColPos, maxColPos);
                if(ret==false && _rangePos>0)
                {
                    _rangePos--;
                    var a = _ranges[_rangePos];

                    _startRow = a.FromRow;
                    _startCol = a.FromCol;
                    _endRow = a.ToRow;
                    _endCol = a.ToCol;
                    _colPos = maxColPos;
                    _row = maxRow;
                    return Previous();
                }
                return ret;
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