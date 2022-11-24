using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml;
using System;
using System.Collections.Generic;


namespace OfficeOpenXml.Core.CellStore
{
    /// <summary>
    /// This class stores ranges to keep track if they have been accessed before.
    /// Merge will add the range and return any part not added before. 
    /// </summary>
    internal class RangeDictionary
    {
        internal Dictionary<int, List<long>> _addresses = new Dictionary<int, List<long>>();
        internal bool Merge(long row, int col)
        {
            long rowSpan = ((row-1) << 20) | (row-1);
            if (!_addresses.TryGetValue(col, out List<long> rows))
            {
                rows = new List<long>();
                _addresses.Add(col, rows);
                rows.Add(rowSpan);
                return true;
            }
            var ix = rows.BinarySearch(rowSpan);
            if (ix < 0)
            {
                ix = ~ix;
                if (ix > 0) ix--;
                var fromRow = (int)(rows[ix] >> 20) + 1;
                var toRow = (int)(rows[ix] & 0xFFFFF) + 1;
                if(row >= fromRow && row <= toRow)
                {
                    return false;
                }
                else
                {
                    if(fromRow-1==row)
                    {
                        rows[ix] = ((row - 1) << 20) | ((long)toRow - 1);
                    }
                    else if(toRow+1==row)
                    {
                        rows[ix] = (((long)fromRow - 1) << 20) | (row - 1);
                    }
                    else
                    {
                        rows[ix] = ((row - 1) << 20) | (row - 1);
                    }
                    return true;
                }
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Merge the cell into the existing data and returns the ranges added.
        /// </summary>
        /// <param name="newAddress"></param>
        /// <returns></returns>
        internal bool Merge(ref FormulaRangeAddress newAddress)
        {
            var spillRanges = new List<long>();
            byte isAdded = 0;
            for (int c = newAddress.FromCol; c <= newAddress.ToCol; c++)
            {
                var rowSpan = (((long)newAddress.FromRow - 1) << 20) | ((long)newAddress.ToRow - 1);
                if (!_addresses.TryGetValue(c, out List<long> rows))
                {
                    rows = new List<long>();
                    rows.Add(rowSpan);
                    spillRanges.Add(rowSpan);
                    _addresses.Add(c, rows);
                    isAdded = 1;
                    continue;
                }
                var ix = rows.BinarySearch(rowSpan);
                if (ix < 0)
                {
                    ix = ~ix;
                    if (ix > 0) ix--;
                    isAdded |= VerifyAndAdd(newAddress, rowSpan, rows, ix, spillRanges);
                    //if (isAdded == false && ++ix < rows.Count)
                    //{
                    //    isAdded = VerifyAndAdd(newAddress, rowSpan, rows, ix, spillRanges);
                    //}
                    MergeWithNext(rows, ix);
                }
                else
                {
                    spillRanges.Add(-1);
                }
            }
            if (isAdded == 1)
            {
                GetSpillRanges(spillRanges, ref newAddress);
            }
            return isAdded != 0;
        }

        private void GetSpillRanges(List<long> spillRanges, ref FormulaRangeAddress address)
        {
            int fromRow, toRow, fromCol, toCol;
            fromRow = toRow = fromCol = toCol = -1;
            var col = address.FromCol;
            bool hasGap = false;
            foreach (var r in spillRanges)
            {
                if (r < -1)
                {
                    return;
                }
                else if (r == -1)
                {
                    if (fromRow > 0)
                    {
                        hasGap = true;
                    }
                }
                else
                {

                    var fr = (int)(r >> 20) + 1;
                    var tr = (int)(r & 0xFFFFF) + 1;
                    if (fromRow == -1)
                    {
                        fromRow = fr;
                        toRow = tr;
                        if (fromRow > 0)
                        {
                            fromCol = toCol = col;
                        }
                    }
                    else
                    {
                        if (fromRow == fr && toRow == tr && hasGap == false)
                        {
                            if (fromCol == 0) fromCol = col;
                            toCol = col;
                        }
                        else
                        {
                            return;
                        }
                    }
                }
                col++;
            }
            address.FromRow = fromRow;
            address.ToRow = toRow;
            address.FromCol = fromCol;
            address.ToCol = toCol;
        }

        private static void MergeWithNext(List<long> rows, int ix)
        {
            do
            {
                if (ix + 1 >= rows.Count) break;
                var fromRow1 = (int)(rows[ix] >> 20) + 1;
                var toRow1 = (int)(rows[ix] & 0xFFFFF) + 1;
                var fromRow2 = (int)(rows[ix + 1] >> 20) + 1;
                var toRow2 = (int)(rows[ix + 1] & 0xFFFFF) + 1;
                if (toRow1 + 1 >= fromRow2)
                {
                    rows[ix] = ((fromRow1 - 1) << 20) | (toRow2 - 1);
                    rows.Remove(rows[ix + 1]);
                }
                else
                {
                    break;
                }
            }
            while (true);
        }

        internal bool Exists(int row, int col)
        {
            if (_addresses.TryGetValue(col, out List<long> rows))
            {
                long rowSpan = ((row - 1) << 20) | (row - 1);
                var ix = rows.BinarySearch(rowSpan);
                if (ix < 0)
                {
                    ix = ~ix;
                    if (ix > 0) ix--;
                    var fromRow = (int)(rows[ix] >> 20) + 1;
                    var toRow = (int)(rows[ix] & 0xFFFFF) + 1;
                    if (row >= fromRow && row <= toRow)
                    {
                        return true;
                    }
                }
                else
                {
                    return true;
                }
            }
            return false;
        }

        private static byte VerifyAndAdd(FormulaRangeAddress newAddress, long rowSpan, List<long> rows, int ix, List<long> spillRanges)
        {
            var fromRow = (int)(rows[ix] >> 20) + 1;
            var toRow = (int)(rows[ix] & 0xFFFFF) + 1;
            if (newAddress.FromRow > toRow)
            {
                if (newAddress.FromRow - 1 == toRow) //Next to each other: Merge
                {
                    rows[ix] = ((long)fromRow - 1 << 20) | (long)(newAddress.ToRow - 1);
                }
                else
                {
                    rows.Insert(ix + 1, rowSpan);
                }
                spillRanges.Add(rowSpan);
                return 1;
            }
            else if (newAddress.ToRow < fromRow)
            {
                if (newAddress.ToRow + 1 == fromRow)   //Next to each other: Merge
                {
                    rows[ix] = ((long)newAddress.FromRow - 1 << 20) | ((long)toRow - 1);
                }
                else
                {
                    rows.Insert(ix, rowSpan);
                }
                spillRanges.Add(rowSpan);
                return 1;
            }
            else
            {
                if (newAddress.FromRow >= fromRow && newAddress.ToRow <= toRow) //Within, 
                {
                    spillRanges.Add(-1);
                }
                else
                {

                    if (newAddress.FromRow < fromRow && newAddress.ToRow <= toRow)
                    {
                        spillRanges.Add(((newAddress.FromRow - 1) << 20) | (fromRow - 2));
                        rows[ix] = (((long)newAddress.FromRow - 1) << 20) | ((long)toRow - 1);
                    }
                    if (newAddress.FromRow >= fromRow && newAddress.ToRow > toRow)
                    {
                        if (newAddress.FromRow < fromRow && newAddress.ToRow <= toRow)
                        {
                            spillRanges[spillRanges.Count - 1] = -2;    //Partial address, return the full address at the end.
                        }
                        else
                        {
                            spillRanges.Add((toRow << 20) | (newAddress.ToRow - 1));
                            rows[ix] = (((long)fromRow - 1) << 20) | ((long)newAddress.ToRow - 1);
                        }
                    }
                    if (newAddress.FromRow < fromRow && newAddress.ToRow > toRow)
                    {
                        spillRanges.Add(-2);
                        rows[ix] = (((long)newAddress.FromRow - 1) << 20) | ((long)newAddress.ToRow - 1);
                    }
                    return 1;
                }
            }

            return 0;
        }
    }
}
