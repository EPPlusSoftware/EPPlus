using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Core.CellStore
{
    /// <summary>
    /// This class stores ranges to keep track if they have been accessed before and adds a reference to <see cref="RangeDictionary{T}"/>.
    /// <typeparamref name="T"/>
    /// </summary>
    internal class RangeDictionary<T>
    {
        internal struct RangeItem : IComparable<RangeItem>
        {
            public RangeItem(long rowSpan)
            {
                RowSpan = rowSpan;
                Value = default;
            }
            public RangeItem(long rowSpan, T value)
            {
                RowSpan= rowSpan;
                Value= value;
            }
            internal long RowSpan;
            internal T Value;

            public int CompareTo(RangeItem other)
            {
                return RowSpan.CompareTo(other.RowSpan);
            }
            public override string ToString()
            {
                var fr = (int)(RowSpan >> 20) + 1;
                var tr = (int)(RowSpan & 0xFFFFF) + 1;
                return $"{fr} - {tr}";
            }
        }
        internal Dictionary<int, List<RangeItem>> _addresses = new Dictionary<int, List<RangeItem>>();
        private bool _extendValuesToInsertedColumn = true;
        internal bool Exists(int fromRow, int fromCol, int toRow, int toCol)
        {
            for (int c = fromCol; c <= toCol; c++)
            {
                var rowSpan = (((long)fromRow - 1) << 20) | ((long)toRow - 1);
                var ri = new RangeItem(rowSpan, default);
                if (_addresses.TryGetValue(c, out List<RangeItem> rows))
                {
                    var ix = rows.BinarySearch(ri);
                    if(ix >= 0)
                    {
                        return true;
                    }
                    else if(rows.Count > 0)
                    {
                        ix = ~ix;
                        if (ix < rows.Count && ExistsInSpan(fromRow, toRow, rows[ix].RowSpan))
                        {
                            return true;
                        }
                        else if(--ix < rows.Count && ix >= 0)
                        {   
                            return ExistsInSpan(fromRow, toRow, rows[ix].RowSpan);
                        }
                    }
                }
            }
            return false;
        }
        internal bool Exists(int row, int col)
        {
            if (_addresses.TryGetValue(col, out List<RangeItem> rows))
            {
                long rowSpan = ((row - 1) << 20) | (row - 1);
                var ri = new RangeItem(rowSpan, default);
                var ix = rows.BinarySearch(ri);
                if (ix < 0)
                {
                    ix = ~ix;
                    if (ix < rows.Count)
                    {
                        if(ExistsInSpan(row, row, rows[ix].RowSpan))
                        {
                            return true;
                        }
                    }
                    if (ix > 0 && --ix < rows.Count)
                    {
                        return ExistsInSpan(row, row, rows[ix].RowSpan);
                    }
                }
                else
                {
                    return true;
                }
            }
            return false;
        }
        internal T this[int row, int column]
        {
            get
            {
                if (_addresses.TryGetValue(column, out List<RangeItem> rows))
                {
                    long rowSpan = ((row - 1) << 20) | (row - 1);
                    var ri = new RangeItem(rowSpan, default);
                    var ix = rows.BinarySearch(ri);
                    if (ix < 0)
                    {
                        ix = ~ix;
                        if (ix < rows.Count)
                        {
                            if (ExistsInSpan(row, row, rows[ix].RowSpan))
                            {
                                return rows[ix].Value;
                            }
                        }
                        if (--ix < rows.Count && ix >= 0)
                        {
                            if(ExistsInSpan(row, row, rows[ix].RowSpan))
                            {
                                return rows[ix].Value;
                            }
                        }
                    }
                    else
                    {
                        return rows[ix].Value;
                    }
                }
                return default;
            }
        }
        internal List<T> GetValuesFromRange(int fromRow, int fromCol, int toRow, int toCol)
        {
            var hs = new HashSet<T>();
            long rowSpan = ((fromRow - 1) << 20) | (fromRow - 1);
            var searchItem = new RangeItem(rowSpan, default);
            var minCol = _addresses.Keys.Min();
            var maxCol = _addresses.Keys.Max();
            fromCol = fromCol < minCol ? minCol : fromCol;
            for (int col = fromCol; col <= toCol; col++)
            {
                if (col > maxCol) break;
                if (_addresses.TryGetValue(col, out List<RangeItem> rows))
                {
                    var ix = rows.BinarySearch(searchItem);
                    if (ix < 0)
                    {
                        ix = ~ix;
                        if (ix > 0) ix--;
                    }
                    while (ix < rows.Count)
                    {
                        var ri = rows[ix];
                        var fr = (int)(ri.RowSpan >> 20) + 1;
                        var tr = (int)(ri.RowSpan & 0xFFFFF) + 1;
                        if (tr < fromRow)
                        {
                            ix++;
                            continue;
                        }
                        if (fromRow <= tr && toRow >= fr)
                        {
                            if (!hs.Contains(ri.Value))
                            {
                                hs.Add(ri.Value);
                            }
                            ix++;
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }
            return hs.ToList();
        }
        internal void Merge(int fromRow, int fromCol, int toRow, int toCol, T value)
        {
            for (int c = fromCol; c <= toCol; c++)
            {
                MergeRowSpan(c, fromRow, toRow, value);
            }
        }
        internal void Add(int fromRow, int fromCol, int toRow, int toCol, T value)
        {
            if (Exists(fromRow, fromCol, toRow, toCol))
            {
                throw (new InvalidOperationException($"Range already starting from range {ExcelCellBase.GetAddress(fromRow, fromCol, toRow, toCol)}"));
            }
            for (int c = fromCol; c <= toCol; c++)
            {
                AddRowSpan(c, fromRow, toRow, value);
            }
        }
        internal void Add(int row, int col, T value)
        {
            if (Exists(row, col))
            {
                throw (new InvalidOperationException($"Range already starting from cell {ExcelCellBase.GetAddress(row, col)}"));
            }
            AddRowSpan(col, row,row, value);
        }
        internal void InsertRow(int fromRow, int noRows, int fromCol = 1, int toCol = ExcelPackage.MaxColumns)
        {
            long rowSpan = ((fromRow - 1) << 20) | (fromRow - 1);
            foreach (var c in _addresses.Keys)
            {
                if (c >= fromCol && c <= toCol)
                {
                    var rows = _addresses[c];
                    var ri = new RangeItem(rowSpan, default);
                    var ix = rows.BinarySearch(ri);
                    if (ix < 0)
                    {
                        ix = ~ix;
                        if (ix > 0) ix--;
                    }

                    if (ix < rows.Count)
                    {
                        ri = rows[ix];
                        var fr = (int)(ri.RowSpan >> 20) + 1;
                        var tr = (int)(ri.RowSpan & 0xFFFFF) + 1;

                        if (tr >= fromRow)
                        {
                            if (fr >= fromRow)
                            {
                                ri.RowSpan = ((fr + noRows - 1) << 20) | (tr + noRows - 1);
                            }
                            else
                            {
                                ri.RowSpan = ((fr - 1) << 20) | (tr + noRows - 1);
                            }
                            rows[ix] = ri;
                        }
                    }
                    var add = (noRows << 20) | (noRows);
                    for (int i = ix + 1; i < rows.Count; i++)
                    {
                        rows[i] = new RangeItem(rows[i].RowSpan + add, rows[i].Value);
                    }
                }
            }
        }

        /// <summary>
        /// Returns empty array if no result because fromRow, toRow covers entire spane
        /// Returns rangeItem with rowspan -1 if the item does not exist within fromRow ToRow
        /// </summary>
        /// <param name="item"></param>
        /// <param name="fromRow"></param>
        /// <param name="toRow"></param>
        /// <returns></returns>
        internal RangeItem[] SplitRangeItem(RangeItem item, int fromRow, int toRow)
        {
            if(ExistsInSpan(fromRow, toRow, item.RowSpan))
            {
                var fromRowRangeItem = (int)(item.RowSpan >> 20) + 1;
                var toRowRangeItem = (int)(item.RowSpan & 0xFFFFF) + 1;

                var topItem = new RangeItem();
                var botItem = new RangeItem();

                long clearedSpan = ((fromRow - 1) << 20) | (toRow - 1);
                var testItem = new RangeItem(clearedSpan);

                var comparedInt = item.CompareTo(testItem);

                List<RangeItem> result = new List<RangeItem>();

                if (fromRow > fromRowRangeItem)
                {
                    int rangeItemOffset = 1;
                    if (fromRow == toRowRangeItem)
                    {
                        rangeItemOffset += 1;
                    }

                    long topSpan = ((fromRowRangeItem - rangeItemOffset) << 20) | (fromRow - 2);
                    topItem.RowSpan = topSpan;
                    topItem.Value = item.Value;
                    result.Add(topItem);
                }

                if (toRow < toRowRangeItem)
                {
                    if (toRow == toRowRangeItem)
                    {
                        toRow += 1;
                    }

                    long endSpan;
                    endSpan = ((toRow) << 20) | (toRowRangeItem - 1);
                    botItem.RowSpan = endSpan;
                    botItem.Value = item.Value;
                    result.Add(botItem);
                }

                return result.ToArray();
            }
            return [new RangeItem(-1L)];
        }

        internal void ClearRows(int fromRow, int noRows, int fromCol = 1, int toCol = ExcelPackage.MaxColumns)
        {
            var toRow = fromRow + noRows - 1;
            //A sheet has a maximum of 65,534 dataValidations. Shifting by 20 more than enough.
            long rowSpan = ((fromRow - 1) << 20) | (fromRow - 1);

            foreach (var c in _addresses.Keys)
            {
                if (c >= fromCol && c <= toCol)
                {
                    var rows = _addresses[c];

                    var ri = new RangeItem(rowSpan);

                    var rowStartIndex = rows.BinarySearch(ri);
                    if (rowStartIndex < 0)
                    {
                        rowStartIndex = ~rowStartIndex;
                        if (rowStartIndex > 0) rowStartIndex--;
                    }

                    for (int i= rowStartIndex; i < rows.Count; i++)
                    {
                        var newItems = SplitRangeItem(rows[i], fromRow, toRow);

                        if(newItems.Length == 0)
                        {
                            rows.Remove(rows[i]);
                            i--;
                            continue;
                        }

                        if (newItems[0].RowSpan == -1)
                        {
                            continue;
                        }

                        rows[i] = newItems[0];
                        if(newItems.Length > 1)
                        {
                            i++;
                            rows.Insert(i, newItems[1]);
                        }
                    }
                }
            }
        }


        internal void DeleteRow(int fromRow, int noRows, int fromCol = 1, int toCol = ExcelPackage.MaxColumns)
        {
            long rowSpan = ((fromRow - 1) << 20) | (fromRow - 1);
            var toRow = fromRow + noRows - 1;
            foreach (var c in _addresses.Keys)
            {
                if (c >= fromCol && c <= toCol)
                {
                    var rows = _addresses[c];
                    var ri = new RangeItem(rowSpan);
                    var rowStartIndex = rows.BinarySearch(ri);
                    if (rowStartIndex < 0)
                    {
                        rowStartIndex = ~rowStartIndex;
                        if (rowStartIndex > 0) rowStartIndex--;
                    }

                    var delete = (noRows << 20) | (noRows);
                    long rowSpanTest = ((fromRow - 1) << 20) | (toRow - 1);
                    var riTest = new RangeItem(rowSpanTest);
                    var deleteTest = new RangeItem(delete);

                    for (int i = rowStartIndex; i < rows.Count; i++)
                    {
                        ri = rows[i];
                        var fromRowRangeItem = (int)(ri.RowSpan >> 20) + 1;
                        var toRowRangeItem = (int)(ri.RowSpan & 0xFFFFF) + 1;

                        bool startAboveOrOnRangeItem = fromRowRangeItem >= fromRow;
                        bool fromRowAboveOrOnEndOfRangeItem = toRowRangeItem >= fromRow;

                        if (startAboveOrOnRangeItem)
                        {
                            if (fromRowRangeItem >= fromRow && toRowRangeItem <= toRow)
                            {
                                rows.RemoveAt(i--);
                                continue;
                            }
                            else if (fromRowRangeItem >= toRow)
                            {

                                fromRowRangeItem = Math.Max(fromRow, fromRowRangeItem - noRows);
                                toRowRangeItem = Math.Max(fromRow, toRowRangeItem - noRows);
                            }
                            else
                            {
                                fromRowRangeItem = Math.Max(fromRow, fromRowRangeItem - noRows);
                                toRowRangeItem = Math.Max(fromRow, toRowRangeItem - noRows);
                            }
                        }
                        else if (fromRowAboveOrOnEndOfRangeItem)
                        {
                            toRowRangeItem = Math.Max(fromRow, toRowRangeItem - noRows);
                        }

                        ri.RowSpan = ((fromRowRangeItem - 1) << 20) | (toRowRangeItem - 1);
                        rows[i] = ri;
                    }
                }
            }
        }
        internal void InsertColumn(int fromCol, int noCols, int fromRow = 1, int toRow = ExcelPackage.MaxRows)
        {
            //Full column
            if (fromRow <= 1 && toRow >= ExcelPackage.MaxRows)
            {
                AddFullColumn(fromCol, noCols);
            }
            else
            {
                InsertPartialColumn(fromCol, noCols, fromRow, toRow);
            }
            if (_extendValuesToInsertedColumn)
            {
                ExtendValues(fromCol - 1, fromCol + noCols, fromRow, toRow);
            }
        }
        private void ExtendValues(int fromCol, int toCol, int fromRow, int toRow)
        {
            if(_addresses.ContainsKey(fromCol) && _addresses.ContainsKey(toCol))
            {
                var toColumn = _addresses[toCol];
                foreach(var item in _addresses[fromCol])
                {
                    var pos = toColumn.BinarySearch(item);
                    if(pos < 0)
                    {
                        pos = ~pos;
                    }
                    while (pos>=0 && pos < toColumn.Count)
                    {
                        var ri = toColumn[pos];
                        var fr = (int)(ri.RowSpan >> 20) + 1;
                        var tr = (int)(ri.RowSpan & 0xFFFFF) + 1;
                        if (tr < fromRow || fr > toRow) break;
                        GetIntersect(item, toColumn[pos], out fr, out tr);
                        if (fr >= 0)
                        {
                            fr = Math.Max(fr, fromRow);
                            tr = Math.Min(tr, toRow);
                            Add(fr, fromCol + 1, tr, toCol - 1, item.Value);
                        }
                        pos++;                        
                    }
                
                }
            }
        }

        private void GetIntersect(RangeItem itemFirst, RangeItem itemLast, out int fr, out int tr)
        {
            if (itemFirst.Value.Equals(itemLast.Value) == false)
            {
                fr = -1;
                tr = -1;
                return;
            }
            var fr1 = (int)(itemFirst.RowSpan >> 20) + 1;
            var tr1 = (int)(itemFirst.RowSpan & 0xFFFFF) + 1;

            var fr2 = (int)(itemLast.RowSpan >> 20) + 1;
            var tr2 = (int)(itemLast.RowSpan & 0xFFFFF) + 1;

            fr=Math.Max(fr1, fr2);
            tr=Math.Min(tr1, tr2);
        }

        internal void DeleteColumn(int fromCol, int noCols, int fromRow = 1, int toRow = ExcelPackage.MaxRows)
        {
            //Full column
            if (fromRow <= 1 && toRow >= ExcelPackage.MaxRows)
            {
                DeleteFullColumn(fromCol, noCols);
            }
            else
            {
                DeletePartialColumn(fromCol, noCols, fromRow, toRow);
            }
        }

        private void DeletePartialColumn(int fromCol, int noCols, int fromRow, int toRow)
        {
            var cols = GetColumnKeys().OrderBy(x=>x);
            var toCol = fromCol + noCols - 1;
            foreach (var colNo in cols)
            {
                if (colNo >= fromCol)
                {
                    if(colNo > toCol)
                    {
                        MoveDataToColumn(colNo, noCols, fromRow, toRow);
                    }
                    DeleteRowsInColumn(colNo, fromRow, toRow);
                }
            }
        }

        private void MoveDataToColumn(int colNo, int noCols, int fromRow, int toRow)
        {
            var destColNo = colNo - noCols;
            if (_addresses.TryGetValue(colNo, out List<RangeItem> sourceCol))
            {
                for (int i = 0; i < sourceCol.Count; i++)
                {
                    var ri = sourceCol[i];
                    var fr = (int)(ri.RowSpan >> 20) + 1;
                    var tr = (int)(ri.RowSpan & 0xFFFFF) + 1;

                    if (fr >= fromRow && tr <= toRow)
                    {
                        Add(fr, destColNo, tr, destColNo, ri.Value);
                    }
                    else if (fr <= fromRow && tr >= fromRow)
                    {
                        Add(fromRow, destColNo, Math.Min(toRow, tr), destColNo, ri.Value);
                    }
                }
            }
        }

        private void DeleteRowsInColumn(int colNo, int fromRow, int toRow)
        {
            var deleteCol = _addresses[colNo];

            for (int i = 0; i < deleteCol.Count; i++)
            {
                var ri = deleteCol[i];
                var fr = (int)(ri.RowSpan >> 20) + 1;
                var tr = (int)(ri.RowSpan & 0xFFFFF) + 1;

                if (fr >= fromRow && tr <= toRow)
                {
                    var rows = tr - fr + 1;
                    DeleteRow(fromRow, tr, colNo, colNo);
                    i--;
                }
                else if (tr >= fromRow)
                {
                    var ntr = fromRow - 1;
                    ri.RowSpan = ri.RowSpan = ((fr - 1) << 20) | (ntr - 1);
                    deleteCol[i] = ri;

                    if (toRow < tr)
                    {
                        Add(toRow + 1, colNo, tr, colNo, ri.Value);
                        i++;
                    }
                }
            }
        }

        private void InsertPartialColumn(int fromCol, int noCols, int fromRow, int toRow)
        {
            var cols = GetColumnKeys();
            foreach (var colNo in cols.OrderByDescending(x => x))
            {
                if (colNo >= fromCol)
                {
                    var sourceCol = _addresses[colNo];
                    for(int i=0; i < sourceCol.Count;i++)
                    {
                        var ri = sourceCol[i];                        
                        var fr = (int)(ri.RowSpan >> 20) + 1;
                        var tr = (int)(ri.RowSpan & 0xFFFFF) + 1;

                        if(fr>=fromRow && tr<=toRow)
                        {
                            var rows = tr - fr + 1;
                            DeleteRow(fromRow, tr, colNo, colNo);
                            Add(fr, colNo + noCols, tr, colNo + noCols, ri.Value);
                            i--;
                        }
                        else if(tr>=fromRow)
                        {
                            var ntr = fromRow-1;
                            ri.RowSpan = ri.RowSpan = ((fr - 1) << 20) | (ntr - 1);
                            sourceCol[i] = ri;

                            Add(fromRow, colNo + noCols, toRow, colNo + noCols, ri.Value);
                            if(toRow<tr)
                            {
                                Add(toRow+1, colNo, tr, colNo, ri.Value);
                                i++;
                            }

                        }
                    }
                }
            }
        }

        private void DeleteFullColumn(int fromCol, int noCols)
        {
            var cols = GetColumnKeys();

            foreach (var key in cols.OrderBy(x => x))
            {
                if (key >= fromCol)
                {
                    if (key < fromCol + noCols)
                    {
                        _addresses.Remove(key);
                    }
                    else
                    {
                        var col = _addresses[key];
                        _addresses.Remove(key);
                        _addresses.Add(key - noCols, col);
                    }
                }
            }
        }

        private void AddFullColumn(int fromCol, int noCols)
        {
            var cols = GetColumnKeys();

            foreach (var key in cols.OrderByDescending(x => x))
            {
                if (key >= fromCol)
                {
                    var col = _addresses[key];
                    _addresses.Remove(key);
                    _addresses.Add(key + noCols, col);
                }
            }
        }

        private List<int> GetColumnKeys()
        {
            var cols = new List<int>();
            foreach (var key in _addresses.Keys)
            {
                cols.Add(key);
            }

            return cols;
        }

        private static bool ExistsInSpan(int fromRow, int toRow, long r)
        {
            var fr = (int)(r >> 20) + 1;
            var tr = (int)(r & 0xFFFFF) + 1;
            return fr <= toRow && tr >= fromRow;
        }
        private void AddRowSpan(int col, int fromRow, int toRow, T value)
        {
            List<RangeItem> rows;
            var rowSpan = ((long)(fromRow - 1) << 20) | (long)(toRow - 1);
            if (_addresses.TryGetValue(col, out rows) == false)
            {
                rows = new List<RangeItem>(64);
                _addresses.Add(col, rows);
            }
            if (rows.Count == 0)
            {
                rows.Add(new RangeItem(rowSpan, value));
                return;
            }
            var ri = new RangeItem(rowSpan, value);
            var ix = rows.BinarySearch(ri);
            if (ix < 0)
            {
                ix = ~ix;
                if (ix < rows.Count)
                {
                    rows.Insert(ix, ri);
                }
                else
                {
                    rows.Add(ri);
                }
            }
        }
        private void MergeRowSpan(int col, int fromRow, int toRow, T value)
        {
            List<RangeItem> rows;
            var rowSpan = ((long)(fromRow - 1) << 20) | (long)(toRow - 1);
            if (_addresses.TryGetValue(col, out rows) == false)
            {
                rows = new List<RangeItem>(64);
                _addresses.Add(col, rows);
            }
            if (rows.Count == 0)
            {
                rows.Add(new RangeItem(rowSpan, value));
                return;
            }
            var ri = new RangeItem(rowSpan, value);
            var ix = rows.BinarySearch(ri);
            if (ix < 0)
            {
                ix = ~ix;
                if (ix > 0) ix--;
                if (ix < rows.Count)
                {
                    int fr, tr = -1;
                    while (rows.Count > ix)
                    {
                        var rs = rows[ix];
                        fr = (int)(rs.RowSpan >> 20) + 1;
                        tr = (int)(rs.RowSpan & 0xFFFFF) + 1;
                        if (fr <= fromRow && tr >= toRow) break; //Inside, exit
                        if (fr > toRow)
                        {
                            rows.Insert(ix, new RangeItem(rowSpan, value));
                            ix++;
                            break;
                        }
                        else if (fromRow < fr)
                        {
                            rowSpan = ((long)(fromRow - 1) << 20) | (long)(fr - 2);
                            rows.Insert(ix, new RangeItem(rowSpan, value));
                            ix++;
                            fromRow = tr + 1;
                        }
                        ix++;
                    }
                    if (tr < toRow)
                    {
                        tr = tr > fromRow - 1 ? tr : fromRow - 1;
                        rowSpan = ((long)(tr) << 20) | (long)(toRow - 1);
                        rows.Insert(ix, new RangeItem(rowSpan, value));
                    }
                }
                else
                {
                    rows.Add(ri);
                }
            }
        }

    }
}
