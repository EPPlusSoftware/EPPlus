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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils
{
    internal static class LookupBinarySearch
    {
        private static int SearchAsc(object s, IRangeInfo lookupRange, IComparer<object> comparer, LookupRangeDirection? direction = null)
        {
            var nRows = lookupRange.Size.NumberOfRows;
            var nCols = lookupRange.Size.NumberOfCols;
            if (nRows == 0 && nCols == 0) return -1;
            int low = 0, high = nRows > nCols ? nRows : nCols, mid;
            if(direction.HasValue)
            {
                high = direction.Value == LookupRangeDirection.Vertical ? nRows : nCols;
            }

            while (low <= high)
            {
                mid = low + high >> 1;

                var col = nRows > nCols ? 0 : mid;
                var row = nRows > nCols ? mid : 0;
                if (direction.HasValue)
                {
                    
                    col = direction.Value == LookupRangeDirection.Vertical ? 0 : mid;
                    row = direction.Value == LookupRangeDirection.Vertical ? mid : 0;
                }
                
                var val = lookupRange.GetOffset(row, col);

                var result = comparer.Compare(s, val);

                if (result < 0)
                    high = mid - 1;

                else if (result > 0)
                    low = mid + 1;

                else
                    return mid;
            }
            return ~low;
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

        internal static int GetMatchIndex(int ix, IRangeInfo returnArray, LookupMatchMode matchMode, bool asc)
        {
            var result = ix < 0 ? ~ix : ix;
            if (matchMode == LookupMatchMode.ExactMatchReturnNextSmaller)
            {
                result = result - 1;
            }
            else if (matchMode == LookupMatchMode.ExactMatchReturnNextLarger)
            {
                var adjustment = asc ? 0 : -1;
                var max = returnArray.Size.NumberOfRows > returnArray.Size.NumberOfCols ?
                    returnArray.Size.NumberOfRows : returnArray.Size.NumberOfCols;
                result = result >= max ? result : result + adjustment;
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
