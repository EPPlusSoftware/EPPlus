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

using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils
{
    internal static class LookupBinarySearch
    {
        private static int SearchAsc(object s, IRangeInfo lookupRange, IComparer<object> comparer, LookupRangeDirection? direction = null)
        {
            var valueSubRange = lookupRange.GetRangeInfoByValue();
            var subrangeAdjustment = valueSubRange.Address.FromRow > 0 ? (valueSubRange.Address.FromRow - lookupRange.Address.FromRow) : 0;

            var searchRange = LookupRangeReader.GetLookupRange(valueSubRange, ref direction);

            var nRows = direction == LookupRangeDirection.Vertical ? searchRange.Count : 1;
            var nCols = direction == LookupRangeDirection.Horizontal ? searchRange.Count : 1;

            if (nRows == 0 && nCols == 0) return -1;
            int low = 0, high = nCols > nRows ? nCols - 1 : nRows - 1, mid;

            if (direction.HasValue)
            {
                high = direction.Value == LookupRangeDirection.Vertical ? nRows : nCols;
                high = high - 1;
            }

            while (low <= high)
            {
                mid = low + high >> 1;

                var searchRangeCell = searchRange[mid];
                var result = comparer.Compare(s, searchRangeCell.Value);
                if (result < 0)
                {
                    if (mid == 0 && high > 0) //First item can be a header, so check the next item for equality as Excel does..
                    {
                        searchRangeCell = searchRange[mid + 1];
                        result = comparer.Compare(s, searchRangeCell.Value);
                        if (result == 0) return 1;
                    }
                    high = mid - 1;
                }
                else if (result > 0)
                {
                    low = mid + 1;
                }
                else
                {
                    return subrangeAdjustment + searchRangeCell.Index;
                }
            }
            if (low < 1)
            {
                return ~(low + subrangeAdjustment);
            }
            return ~(searchRange[low - 1].Index + subrangeAdjustment + 1);
        }

        private static int SearchDesc(object s, IRangeInfo lookupRange, IComparer<object> comparer)
        {
            var nRows = lookupRange.Size.NumberOfRows;
            var nCols = lookupRange.Size.NumberOfCols;
            if (nRows == 0 && nCols == 0) return -1;
            int low = 0, high = nRows > nCols ? nRows : nCols, mid;

            while (high >= low)
            {
                mid = high + low >> 1;

                var col = nRows > nCols ? 0 : mid;
                var row = nRows > nCols ? mid : 0;
                var val = lookupRange.GetOffset(row, col);

                var result = comparer.Compare(s, val);

                if (result < 0)
                    low = mid + 1;

                else if (result > 0)
                    high = mid - 1;

                else
                    return mid;
            }
            return ~low;
        }

        internal static int SearchAsc(object s, LookupSearchItem[] items, IComparer<object> comparer)
        {
            if (items.Length == 0) return -1;
            int low = 0, high = items.Length - 1, mid;

            while (low <= high)
            {
                mid = low + high >> 1;

                var result = comparer.Compare(s, items[mid].Value);

                if (result < 0)
                    high = mid - 1;

                else if (result > 0)
                    low = mid + 1;

                else
                    return mid;
            }
            return ~low;
        }

        internal static int SearchDesc(object s, LookupSearchItem[] items, IComparer<object> comparer)
        {
            if (items.Length == 0) return -1;
            int low = 0, high = items.Length - 1, mid;

            while (high >= low)
            {
                mid = high + low >> 1;

                var result = comparer.Compare(s, items[mid].Value);

                if (result < 0)
                    low = mid + 1;

                else if (result > 0)
                    high = mid - 1;

                else
                    return mid;
            }
            return ~low;
        }

        internal static int BinarySearch(object lookupValue, IRangeInfo lookupRange, bool asc, IComparer<object> comparer, LookupRangeDirection? direction = null)
        {
            return asc ? SearchAsc(lookupValue, lookupRange, comparer, direction) : SearchDesc(lookupValue, lookupRange, comparer);
        }

        /// <summary>
        /// Binary search over the lookup range for approximate match lookups
        /// (e.g. VLOOKUP/HLOOKUP with range_lookup = TRUE).
        /// Leading empty cells are skipped, but inner and trailing empty cells are
        /// kept in their original positions. This mirrors Excel: leading blank rows
        /// (common when a whole-column reference such as B:C is used and the data
        /// starts further down) must not drag the search away from the data, while
        /// blank cells inside the data block affect how the binary search partitions
        /// the range and must be preserved to match Excel's result on such data.
        /// Note that the previous approach of removing all empty cells (via
        /// <see cref="LookupRangeReader.GetLookupRange"/>) shifted the midpoints and
        /// could produce a different result than Excel.
        /// Returns the 0-based offset (row offset for vertical, column offset for
        /// horizontal) into the original range, or a bitwise complement of the
        /// insertion point when no exact match is found (to be resolved by
        /// <see cref="GetMatchIndex(int, IRangeInfo, LookupMatchMode, bool)"/>).
        /// </summary>
        internal static int SearchAscFullRange(object lookupValue, IRangeInfo lookupRange, IComparer<object> comparer, LookupRangeDirection direction)
        {
            var count = direction == LookupRangeDirection.Vertical
                ? lookupRange.Size.NumberOfRows
                : lookupRange.Size.NumberOfCols;
            if (count == 0) return -1;

            //Find the first non-empty cell efficiently. Using GetRangeInfoByValue()
            //jumps directly to the first/last value cell via the cell store instead
            //of scanning cell-by-cell, which is essential for whole-column references
            //such as A:A where 'count' can be ~1,000,000.
            //Only the leading edge is used: the search still runs to the original end
            //of the range so that inner and trailing blanks keep their positions.
            var startOffset = GetFirstValueOffset(lookupRange, direction);
            if (startOffset < 0 || startOffset >= count) return -1;

            int low = startOffset, high = count - 1, mid;
            while (low <= high)
            {
                mid = low + high >> 1;
                var cellValue = GetCellValue(lookupRange, mid, direction);
                var result = comparer.Compare(lookupValue, cellValue);
                if (result < 0)
                {
                    if (mid == startOffset && high > startOffset)
                    {
                        //The first item can be a header that is not comparable to the
                        //lookup value (e.g. a text header above numeric data). Excel
                        //skips such a header, so check the next item for an exact match.
                        var nextValue = GetCellValue(lookupRange, mid + 1, direction);
                        if (comparer.Compare(lookupValue, nextValue) == 0)
                        {
                            return mid + 1;
                        }
                    }
                    high = mid - 1;
                }
                else if (result > 0)
                {
                    low = mid + 1;
                }
                else
                {
                    return mid;
                }
            }
            //If the insertion point is at the start of the data, the lookup value is
            //smaller than every value in the range, which Excel reports as #N/A.
            //Returning -1 here prevents NextSmaller from pointing into the skipped
            //leading blank cells.
            if (low <= startOffset)
            {
                return -1;
            }
            return ~low;
        }

        private static object GetCellValue(IRangeInfo lookupRange, int offset, LookupRangeDirection direction)
        {
            return direction == LookupRangeDirection.Vertical
                ? lookupRange.GetOffset(offset, 0)
                : lookupRange.GetOffset(0, offset);
        }

        /// <summary>
        /// Returns the offset (row offset for vertical, column offset for horizontal)
        /// of the first non-empty cell in the lookup range, relative to the top-left
        /// of the range. Uses <see cref="IRangeInfo.GetRangeInfoByValue"/>, which
        /// locates the first value cell through the cell store rather than scanning
        /// each cell, so this is safe for whole-column/row references. Returns -1 if
        /// the range contains no values.
        /// </summary>
        private static int GetFirstValueOffset(IRangeInfo lookupRange, LookupRangeDirection direction)
        {
            var valueSubRange = lookupRange.GetRangeInfoByValue();
            if (valueSubRange == null) return -1;

            var valueAddress = valueSubRange.Address;
            var rangeAddress = lookupRange.Address;
            if (valueAddress == null || rangeAddress == null) return -1;
            if (valueAddress.FromRow < 0 || valueAddress.FromCol < 0) return -1;

            return direction == LookupRangeDirection.Vertical
                ? valueAddress.FromRow - rangeAddress.FromRow
                : valueAddress.FromCol - rangeAddress.FromCol;
        }

        internal static int GetMaxIndex(IRangeInfo returnArray)
        {
            return returnArray.Size.NumberOfRows > returnArray.Size.NumberOfCols ?
                    returnArray.Size.NumberOfRows : returnArray.Size.NumberOfCols;
        }

        internal static int GetMatchIndex(int ix, IRangeInfo returnArray, LookupMatchMode matchMode, bool asc)
        {
            if (ix > -1) return ix;
            var result = ~ix;
            if (matchMode == LookupMatchMode.ExactMatchReturnNextSmaller)
            {
                result = result - 1;
            }
            else if (matchMode == LookupMatchMode.ExactMatchReturnNextLarger)
            {
                var adjustment = asc ? 0 : -1;
                var max = GetMaxIndex(returnArray);
                result = result >= max ? max : result + adjustment;
            }
            return result;
        }

        internal static int GetMatchIndex(object lookupValue, List<LookupSearchItem> searchItems, LookupSearchMode searchMode, LookupMatchMode matchMode)
        {
            var saf = searchMode == LookupSearchMode.StartingAtFirst;
            var startIx = saf ? 0 : searchItems.Count - 1;
            var endIx = saf ? searchItems.Count - 1 : 0;
            var incrementor = saf ? 1 : -1;

            var comparer = new LookupComparer(matchMode);
            for (var ix = startIx; saf ? ix <= endIx : ix > endIx; ix += incrementor)
            {
                var item = searchItems[ix];
                var cr = comparer.Compare(lookupValue, item.Value);
                if (cr == 0)
                {
                    return item.OriginalIndex;
                }
                else if (cr < 0)
                {
                    if (matchMode == LookupMatchMode.ExactMatchReturnNextSmaller && ix > 0)
                    {
                        return searchItems[ix - 1].OriginalIndex;
                    }
                    else if (matchMode == LookupMatchMode.ExactMatchReturnNextLarger)
                    {
                        return ix < searchItems.Count - 1 ? searchItems[ix + 1].OriginalIndex : searchItems[ix].OriginalIndex;
                    }
                }
            }
            return -1;
        }
    }
}