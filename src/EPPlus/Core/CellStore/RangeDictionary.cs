using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System.Runtime.CompilerServices;
using System.Net;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System.Security.Cryptography;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace OfficeOpenXml.Core.CellStore
{
    /// <summary>
    /// This class stores ranges to keep track if they have been accessed before and adds a reference to <see cref="T"/>.
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
                    if (--ix < rows.Count)
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
        internal void InsertRow(int fromRow, int noRows, int fromCol=1, int toCol=ExcelPackage.MaxColumns)
        {
            long rowSpan = ((fromRow - 1) << 20) | (fromRow - 1);            
            foreach(var c in _addresses.Keys)
            {
                if(c>=fromCol && c <= toCol)
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

                        if(fr>=fromRow)
                        {
                            ri.RowSpan = ((fr + noRows - 1) << 20) | (tr + noRows - 1);
                        }
                        else
                        {
                            ri.RowSpan = ((fr - 1) << 20) | (tr + noRows - 1);
                        }
                        rows[ix] = ri;
                    }
                    var add = (noRows << 20) | (noRows);
                    for (int i=ix+1;i<rows.Count;i++)
                    {
                        rows[i]= new RangeItem(rows[i].RowSpan+add, rows[i].Value);
                    }
                }
            }
        }
        internal void DeleteRow(int fromRow, int noRows, int fromCol = 1, int toCol = ExcelPackage.MaxColumns)
        {
            long rowSpan = ((fromRow - 1) << 20) | (fromRow - 1);
            foreach (var c in _addresses.Keys)
            {
                if (c >= fromCol && c <= toCol)
                {
                    var rows = _addresses[c];
                    var ri = new RangeItem(rowSpan);
                    var ix = rows.BinarySearch(ri);
                    if (ix < 0)
                    {
                        ix = ~ix;
                        if (ix > 0) ix--;
                    }

                    var delete = (noRows << 20) | (noRows);
                    for (int i = ix; i < rows.Count; i++)
                    {
                        ri = rows[i];
                        var fr = (int)(ri.RowSpan >> 20) + 1;
                        var tr = (int)(ri.RowSpan & 0xFFFFF) + 1;

                        if(fr >= fromRow)
                        {
                            if(fr >= fromRow && tr <= fromRow + noRows)
                            { 
                                rows.RemoveAt(ix--);
                                continue;
                            }
                            else if(fr >= fromRow + noRows)
                            {
                                
                                tr -= noRows;
                                fr -= noRows;
                            }
                            else
                            {
                                fr = Math.Max(fromRow, fr - noRows);
                                tr = Math.Max(fromRow, tr - noRows);
                            }
                        }
                        else if(fr+noRows >= fromRow) 
                        {
                            tr = Math.Max(fromRow, tr - noRows);
                        }

                        ri.RowSpan = ((fr - 1) << 20) | (tr - 1);
                        rows[i] = ri;
                    }
                }
            }
        }
        internal void InsertColumn(int fromCol, int noCols, int fromRow = 1, int toRow = ExcelPackage.MaxRows)
        {
            //Full column
            if(fromRow<=1 && toRow >=ExcelPackage.MaxRows)
            {
                AddFullColumn(fromCol, noCols);
            }
            else
            {
                InsertPartialColumn(fromCol, noCols, fromRow, toRow);
            }
            if(_extendValuesToInsertedColumn)
            {
                ExtendValues(fromCol - 1, fromCol+noCols, fromRow, toRow);
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
                        GetItersect(item, toColumn[pos], out fr, out tr);
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

        private void GetItersect(RangeItem itemFirst, RangeItem itemLast, out int fr, out int tr)
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
                        if(fromRow<fr)
                        {
                            rowSpan = ((long)(fromRow - 1) << 20) | (long)(fr - 2);
                            rows.Insert(ix, new RangeItem(rowSpan, value));
                            ix++;
                            fromRow=tr+1;
                        }
                        ix++;
                    }
                    if(tr < toRow)
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
