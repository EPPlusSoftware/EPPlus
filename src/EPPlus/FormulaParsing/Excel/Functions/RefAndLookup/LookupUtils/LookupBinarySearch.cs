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

            //Excel performs a plain binary search over the WHOLE range for approximate
            //match. Empty cells are kept in their original positions (they affect how
            //the search partitions the range), and an empty cell counts as "greater"
            //than the lookup value so the search moves left. We therefore do NOT trim
            //or compact the range in any way; doing so shifts the midpoint sequence and
            //produces results that differ from Excel.
            //
            //GetRangeInfoByValue() is used only to obtain the value edges:
            //- firstValueOffset: lets us apply the header rule and report #N/A when the
            //  lookup value is smaller than every value.
            //- lastValueOffset: bounds the inner blank-skipping scan so a whole-column
            //  reference (A:A, count ~1,000,000) never scans the empty tail. This is a
            //  performance bound only; it does not change the result because everything
            //  past the last value is empty anyway.
            if (!TryGetValueEdges(lookupRange, direction, out var firstValueOffset, out var lastValueOffset))
            {
                return -1;
            }
            if (firstValueOffset < 0 || firstValueOffset >= count) return -1;

            int low = 0, high = count - 1, mid;
            while (low <= high)
            {
                mid = low + high >> 1;

                //If the midpoint is an empty cell it gives no direction information.
                //Look forward (towards higher indexes) to the next non-empty cell and
                //compare against that, mirroring how Excel sees past blank cells.
                //The scan is bounded by the last value position, so a whole-column
                //reference never scans across the empty remainder of the column.
                var probe = mid;
                var scanLimit = high < lastValueOffset ? high : lastValueOffset;
                var cellValue = GetCellValue(lookupRange, probe, direction);
                while (probe < scanLimit && cellValue == null)
                {
                    probe++;
                    cellValue = GetCellValue(lookupRange, probe, direction);
                }
                if (cellValue == null)
                {
                    //Only empty cells from 'mid' up to the scan limit: nothing to match
                    //on the right, so narrow the search downwards.
                    high = mid - 1;
                    continue;
                }

                var result = comparer.Compare(lookupValue, cellValue);
                if (result < 0)
                {
                    if (probe == firstValueOffset && high > probe)
                    {
                        //The first value can be a header that is not comparable to the
                        //lookup value (e.g. a text header above numeric data). Excel
                        //skips such a header, so check the next item for an exact match.
                        var nextValue = GetCellValue(lookupRange, probe + 1, direction);
                        if (nextValue != null && comparer.Compare(lookupValue, nextValue) == 0)
                        {
                            return probe + 1;
                        }
                    }
                    high = probe - 1;
                }
                else if (result > 0)
                {
                    low = probe + 1;
                }
                else
                {
                    return probe;
                }
            }
            //The next-smaller insertion point is low - 1. If that falls before the first
            //value, the lookup value is smaller than every value in the range, which
            //Excel reports as #N/A.
            var matchIndex = low - 1;
            if (matchIndex < firstValueOffset)
            {
                return -1;
            }
            return matchIndex;
        }

        private static object GetCellValue(IRangeInfo lookupRange, int offset, LookupRangeDirection direction)
        {
            return direction == LookupRangeDirection.Vertical
                ? lookupRange.GetOffset(offset, 0)
                : lookupRange.GetOffset(0, offset);
        }

        /// <summary>
        /// Gets the offsets (relative to the top-left of the range) of the first and
        /// last non-empty cell along the search direction. Uses
        /// <see cref="IRangeInfo.GetRangeInfoByValue"/>, which locates the value edges
        /// through the cell store rather than scanning each cell, so this is safe for
        /// whole-column/row references. Returns false if the range contains no values.
        /// </summary>
        private static bool TryGetValueEdges(IRangeInfo lookupRange, LookupRangeDirection direction, out int startOffset, out int lastValueOffset)
        {
            startOffset = -1;
            lastValueOffset = -1;

            var valueSubRange = lookupRange.GetRangeInfoByValue();
            if (valueSubRange == null) return false;

            var valueAddress = valueSubRange.Address;
            var rangeAddress = lookupRange.Address;
            if (valueAddress == null || rangeAddress == null) return false;
            if (valueAddress.FromRow < 0 || valueAddress.FromCol < 0) return false;

            if (direction == LookupRangeDirection.Vertical)
            {
                startOffset = valueAddress.FromRow - rangeAddress.FromRow;
                lastValueOffset = valueAddress.ToRow - rangeAddress.FromRow;
            }
            else
            {
                startOffset = valueAddress.FromCol - rangeAddress.FromCol;
                lastValueOffset = valueAddress.ToCol - rangeAddress.FromCol;
            }
            return true;
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