/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB.
  This software is licensed under PolyForm Noncommercial License 1.0.0
  and may only be used for noncommercial purposes
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/25/2026         EPPlus Software AB       Avoid walking empty cells in conditional formatting ranges
 *************************************************************************************************/
using OfficeOpenXml.Core.CellStore;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.ConditionalFormatting.Rules
{
    /// <summary>
    /// Helpers for reading the values in a conditional formatting range without visiting empty cells.
    /// The range (sqref) of a rule can cover a whole sheet, so walking it cell by cell is not an option.
    /// </summary>
    internal static class CFRangeValues
    {
        /// <summary>
        /// Returns the non-null values in the address.
        /// Sub-addresses are handled in order and each one is returned row by row, column by column,
        /// which is the same order as walking every cell of the address.
        /// Cells covered by more than one sub-address are returned once per sub-address.
        /// </summary>
        internal static List<object> GetValues(ExcelWorksheet ws, ExcelAddressBase address)
        {
            var result = new List<object>();
            var cells = new List<KeyValuePair<long, object>>();
            foreach (var a in address.GetAllAddresses())
            {
                cells.Clear();
                var e = new CellStoreEnumerator<ExcelValue>(ws._values, a._fromRow, a._fromCol, a._toRow, a._toCol);
                while (e.Next())
                {
                    var row = e.Row;
                    var col = e.Column;
                    if (row < a._fromRow || row > a._toRow || col < a._fromCol || col > a._toCol) continue;
                    var v = ws.GetValue(row, col);
                    if (v != null)
                    {
                        cells.Add(new KeyValuePair<long, object>(((long)row << 16) | (uint)col, v));
                    }
                }
                cells.Sort((x, y) => x.Key.CompareTo(y.Key));
                foreach (var c in cells)
                {
                    result.Add(c.Value);
                }
            }
            return result;
        }

        /// <summary>
        /// Total number of cells in the address, including empty cells. Cells covered by more than one sub-address are counted once per sub-address.
        /// </summary>
        internal static long GetCellCount(ExcelAddressBase address)
        {
            long count = 0;
            foreach (var a in address.GetAllAddresses())
            {
                count += (long)a.Rows * a.Columns;
            }
            return count;
        }

        /// <summary>
        /// Returns the number of items in the sorted array that are less than <paramref name="value"/>, using <see cref="double.CompareTo(double)"/> ordering (NaN sorts first).
        /// </summary>
        internal static int CountLessThan(double[] sorted, double value)
        {
            int lo = 0, hi = sorted.Length;
            while (lo < hi)
            {
                int mid = lo + ((hi - lo) >> 1);
                if (sorted[mid].CompareTo(value) < 0)
                {
                    lo = mid + 1;
                }
                else
                {
                    hi = mid;
                }
            }
            return lo;
        }
    }
}
