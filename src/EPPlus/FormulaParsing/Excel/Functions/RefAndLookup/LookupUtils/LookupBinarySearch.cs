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

using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils
{
    internal static class LookupBinarySearch
    {
        private static int SearchAsc(object s, IRangeInfo lookupRange, IComparer<object> comparer, LookupRangeDirection? direction = null)
        {
            //Only look at relevant values. We use this to omit null values above and below actual values
            var valueSubRange = ValueFinder.RangeByValue(lookupRange, out CellStore<object> CsValues);

            List<List<object>> objects = new();

            if(lookupRange.IsInMemoryRange)
            {

            }
            else 
            {
                if(lookupRange.Address.ToExcelAddressBase().IsExternal)
                {
                    var extRange = (EpplusExcelExternalRangeInfo)lookupRange;
                    for(int i =0; i< lookupRange.Size.NumberOfCols; i++)
                    {
                        objects.Add(extRange._externalWs.CellValues._values._columnIndex[i]._values);
                    }
                }
                else
                {
                    //Non-external address
                    var range = (EpplusExcelExternalRangeInfo)lookupRange;
                    for (int i = 0; i < lookupRange.Size.NumberOfCols; i++)
                    {
                        objects.Add(range._externalWs.CellValues._values._columnIndex[i]._values);
                    }
                }
            }
            

            var nRows = valueSubRange.Size.NumberOfRows;
            var nCols = valueSubRange.Size.NumberOfCols;

            if (nRows == 0 && nCols == 0) return -1;
            int low = 0, high = nCols > nRows ? nCols : nRows, mid;

            if (direction.HasValue)
            {
                high = direction.Value == LookupRangeDirection.Vertical ? nRows : nCols;
            }

            while (low <= high)
            {
                mid = low + high >> 1;

                var col = nRows >= nCols ? 0 : mid;
                var row = nRows >= nCols ? mid : 0;
                if (direction.HasValue)
                {
                    col = direction.Value == LookupRangeDirection.Vertical ? 0 : mid;
                    row = direction.Value == LookupRangeDirection.Vertical ? mid : 0;
                }

                //Row and col are 0-based if equal we will be past the last value due to GetOffset
                if (row == nRows || col == nCols)
                {
                    break;
                }

                var val = valueSubRange.GetOffset(row, col);
                if (val == null)
                {
                    var currRow = valueSubRange.Address.FromRow + row;
                    var currCol = valueSubRange.Address.FromCol + col;
                    var exists = CsValues.NextCell(ref currRow, ref currCol);

                    row = currRow - valueSubRange.Address.FromRow;
                    col = currCol - valueSubRange.Address.FromCol;

                    low = mid + 1;
                    continue;
                }
                var result = comparer.Compare(s, val);

                if (result < 0)
                {
                    high = mid - 1;
                }

                else if (result > 0)
                {
                    low = mid + 1;
                }
                else
                {
                    return valueSubRange.Address.FromRow - 1 + mid;
                }
            }

            return ~(low + valueSubRange.Address.FromRow + 1);
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
