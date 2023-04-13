/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/3/2023         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Sorting
{
    internal class SortByImpl
    {
        private class SortInfo
        {
            public SortInfo()
            {
                RangeData = new List<object>();
                CompareData = new List<object>();
            }
            public List<object> RangeData
            {
                get; private set;
            }

            public List<object> CompareData
            {
                get; private set;
            }

            public bool CompareValueWasNull
            {
                get; set;
            }
        }
        public SortByImpl(IRangeInfo sourceRange, List<IRangeInfo> byRanges, List<short> sortOrders, LookupDirection direction)
        {
            _sourceRange= sourceRange;
            _byRanges= byRanges;
            _sortOrders= sortOrders;
            _direction= direction;
        }

        private readonly IRangeInfo _sourceRange;
        private readonly List<IRangeInfo> _byRanges;
        private readonly List<short> _sortOrders;
        private readonly LookupDirection _direction;
        private readonly LookupComparerBase _comparer = new SortByComparer();

        public IRangeInfo Sort()
        {
            var inMemoryRange = new InMemoryRange(_sourceRange.Size);
            if (_direction == LookupDirection.Vertical)
            {
                var rows = GetSortedRows();
                for(var row = 0; row < _sourceRange.Size.NumberOfRows; row++)
                {
                    for(var col = 0; col  < _sourceRange.Size.NumberOfCols; col++)
                    {
                        inMemoryRange.SetValue(row, col, rows[row].RangeData[col]);
                    }
                }
            }
            else
            {
                var cols = GetSortedCols();
                for (var row = 0; row < _sourceRange.Size.NumberOfRows; row++)
                {
                    for (var col = 0; col < _sourceRange.Size.NumberOfCols; col++)
                    {
                        inMemoryRange.SetValue(row, col, cols[col].RangeData[row]);
                    }
                }
            }
            return inMemoryRange;
        }

        private List<SortInfo> GetSortedRows()
        {
            var rows = new List<SortInfo>();
            for(var row = 0; row < _sourceRange.Size.NumberOfRows; row++)
            {
                var cols = new SortInfo();
                for(var col = 0; col  < _sourceRange.Size.NumberOfCols; col++)
                {
                    var v = _sourceRange.GetOffset(row, col);
                    cols.RangeData.Add(v);
                }
                for(var byRange = 0; byRange < _byRanges.Count; byRange++)
                {
                    cols.CompareData.Add(_byRanges[byRange].GetOffset(row, 0));
                }
                rows.Add(cols);
            }
            rows.Sort((a, b) =>
            {
                var compareDataIx = -1;
                while((a.CompareData != null && b.CompareData != null) && compareDataIx < a.CompareData.Count - 1  && compareDataIx < b.CompareData.Count -1 )
                {
                    compareDataIx++;
                    var asi = a.CompareData[compareDataIx];
                    var bsi = b.CompareData[compareDataIx];
                    var cr = _comparer.Compare(asi, bsi, _sortOrders[compareDataIx]);
                    if(cr != 0)
                    {
                        return cr;
                    }
                }
                return 0;
            });
            return rows;
        }

        private List<SortInfo> GetSortedCols()
        {
            var cols = new List<SortInfo>();
            for (var col = 0; col < _sourceRange.Size.NumberOfCols; col++)
            {
                var rows = new SortInfo();
                for (var row = 0; row < _sourceRange.Size.NumberOfRows; row++)
                {
                    var v = _sourceRange.GetOffset(row, col);
                    rows.RangeData.Add(v);
                }
                for (var byRange = 0; byRange < _byRanges.Count; byRange++)
                {
                    rows.CompareData.Add(_byRanges[byRange].GetOffset(0, col));
                }
                cols.Add(rows);
            }
            cols.Sort((a, b) =>
            {
                var compareDataIx = -1;
                while ((a.CompareData != null && b.CompareData != null) && compareDataIx < a.CompareData.Count - 1 && compareDataIx < b.CompareData.Count - 1)
                {
                    compareDataIx++;
                    var cr = _comparer.Compare(a.CompareData[compareDataIx], b.CompareData[compareDataIx], _sortOrders[compareDataIx]);
                    if (cr != 0)
                    {
                        return cr;
                    }
                }
                return 0;
            });
            return cols;
        }
    }
}
