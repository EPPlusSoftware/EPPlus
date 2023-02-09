using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System.Runtime.CompilerServices;
using System.Net;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

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
                    else
                    {
                        ix = ~ix;
                        if (ix < rows.Count)
                        {
                            return ExistsInSpan(fromRow, toRow, rows[ix].RowSpan);
                        }
                        else if(--ix < rows.Count)
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
        private static bool ExistsInSpan(int fromRow, int toRow, long r)
        {
            var fr = (int)(r >> 20) + 1;
            var tr = (int)(r & 0xFFFFF) + 1;
            return fr <= toRow && tr >= fromRow;
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
        private void AddRowSpan(int col, int fromRow, int toRow, T value)
        {
            List<RangeItem> rows;
            long rowSpan = ((fromRow - 1) << 20) | (toRow - 1);
            if (_addresses.TryGetValue(col, out rows) == false)
            {
                rows = new List<RangeItem>();
                _addresses.Add(col, rows);
            }
            if(rows.Count==0)
            {
                rows.Add(new RangeItem(rowSpan, value));
                return;
            }
            var ri = new RangeItem(rowSpan, value);
            var ix = rows.BinarySearch(ri);
            if(ix < 0)
            {
                ix = ~ix;
                if(ix < rows.Count)
                {
                    rows.Insert(ix, ri);
                }
                else
                {
                    rows.Add(new RangeItem(rowSpan, value));
                }
            }
        }
        internal void InsertRow(int fromRow, int noRows, int fromCol=1, int toCol=ExcelPackage.MaxRows)
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
        internal void DeleteRow(int fromRow, int noRows, int fromCol = 1, int toCol = ExcelPackage.MaxRows)
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
                            if(fr >= fromRow + noRows && tr <= fromRow + noRows)
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
    }
}
