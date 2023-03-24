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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.XlookupUtils;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using static OfficeOpenXml.FormulaParsing.Excel.Functions.Math.RoundingHelper;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
            Category = ExcelFunctionCategory.LookupAndReference,
            EPPlusVersion = "6.0",
            IntroducedInExcelVersion = "2016",
            Description = "Searches a range or an array, and then returns the item corresponding to the first match it finds. Will return a VALUE error if the functions returns an array (EPPlus does not support dynamic arrayformulas)", 
            SupportsArrays = true)]
    internal class Xlookup : LookupFunction
    {
        private enum XlookupRangeDirection
        {
            Vertical,
            Horizontal
        }

        private readonly ValueMatcher _valueMatcher = new WildCardValueMatcher();
        private XlookupRangeDirection GetLookupDirection(IRangeInfo lookupRange)
        {
            var result = XlookupRangeDirection.Vertical;
            if(lookupRange.Size.NumberOfCols > 1)
            {
                result = XlookupRangeDirection.Horizontal;
            }
            return result;
        }

        private XLookupMatchMode GetMatchMode(int mm)
        {
            switch (mm)
            {
                case 0:
                    return XLookupMatchMode.ExactMatch;
                case -1:
                    return XLookupMatchMode.ExactMatchReturnNextSmaller;
                case 1:
                    return XLookupMatchMode.ExactMatchReturnNextLarger;
                case 2:
                    return XLookupMatchMode.Wildcard;
                default:
                    throw new ArgumentException("Invalid match mode: " + mm.ToString());
            }
        }

        private XLookupSearchMode GetSearchMode(int sm)
        {
            switch (sm)
            {
                case 1:
                    return XLookupSearchMode.StartingAtFirst;
                case -1:
                    return XLookupSearchMode.ReverseStartingAtLast;
                case 2:
                    return XLookupSearchMode.BinarySearchAscending;
                case -2:
                    return XLookupSearchMode.BinarySearchDescending;
                default:
                    throw new ArgumentException("Invalid search mode: " + sm.ToString());

            }
        }

        private List<XlookupSearchItem> GetSortedArray(List<object> arr)
        {
            var sortedList = new List<XlookupSearchItem>();
            if (arr == null || arr.Count == 0) return sortedList;
            for(var ix = 0; ix < arr.Count; ix++)
            {
                sortedList.Add(new XlookupSearchItem(arr[ix], ix));
            }
            sortedList.Sort((a, b) => CompareObjects(a.Value, b.Value));
            return sortedList;
        }

        private static int CompareObjects(object x1, object y1)
        {
            int ret;
            var isNumX = ConvertUtil.IsNumericOrDate(x1);
            var isNumY = ConvertUtil.IsNumericOrDate(y1);
            if (isNumX && isNumY)   //Numeric Compare
            {
                var d1 = ConvertUtil.GetValueDouble(x1);
                var d2 = ConvertUtil.GetValueDouble(y1);
                if (double.IsNaN(d1))
                {
                    d1 = double.MaxValue;
                }
                if (double.IsNaN(d2))
                {
                    d2 = double.MaxValue;
                }
                ret = d1 < d2 ? -1 : (d1 > d2 ? 1 : 0);
            }
            else if (isNumX == false && isNumY == false)   //String Compare
            {
                var s1 = x1 == null ? "" : x1.ToString();
                var s2 = y1 == null ? "" : y1.ToString();
                ret = string.Compare(s1, s2, StringComparison.CurrentCulture);
            }
            else
            {
                ret = isNumX ? -1 : 1;
            }

            return ret;
        }

        private List<object> GetLookupArray(IRangeInfo lookupRange, XlookupRangeDirection direction)
        {
            var arr = new List<object>();
            if (direction == XlookupRangeDirection.Vertical)
            {
                for (var rowIx = 0; rowIx < lookupRange.Size.NumberOfRows; rowIx++)
                {
                    arr.Add(lookupRange.GetOffset(rowIx, 0));
                }
            }
            else
            {
                for (var colIx = 0; colIx < lookupRange.Size.NumberOfCols; colIx++)
                {
                    arr.Add(lookupRange.GetOffset(0, colIx));
                }
            }
            return arr;
        }

        private int GetMatchIndex(object lookupValue, IRangeInfo lookupRange, XLookupSearchMode searchMode, XLookupMatchMode matchMode)
        {
            var ix = -1;
            var comparer = new XlookupObjectComparer(matchMode);
            if (searchMode == XLookupSearchMode.BinarySearchAscending)
            {
                ix = XLookupBinarySearch.Search(lookupValue, lookupRange, comparer);
            }
            else if (searchMode == XLookupSearchMode.BinarySearchDescending)
            {
                //searchItems.Sort((a, b) => comparer.Compare(b.Value, a.Value));
                ix = XLookupBinarySearch.SearchDesc(lookupValue, lookupRange, comparer);
            }
            return ix;
        }

        private int GetMatchIndex(object lookupValue, List<XlookupSearchItem> searchItems, XLookupSearchMode searchMode, XLookupMatchMode matchMode)
        {
            var saf = searchMode == XLookupSearchMode.StartingAtFirst;
            var startIx = saf ? 0 : searchItems.Count - 1;
            var endIx = saf ? searchItems.Count - 1 : 0;
            var incrementor = saf ? 1 : -1;
            var comparer = new XlookupObjectComparer(matchMode);
            for (var ix = startIx; saf ? ix <= endIx : ix > endIx; ix += incrementor)
            {
                var item = searchItems[ix];
                var cr = comparer.Compare(lookupValue, item.Value);
                if (cr == 0)
                {
                    return item.OriginalIndex;
                }
                else if(cr < 0)
                {
                    if(matchMode == XLookupMatchMode.ExactMatchReturnNextSmaller && ix > 0)
                    {
                        return searchItems[ix - 1].OriginalIndex;
                    }
                    else if(matchMode == XLookupMatchMode.ExactMatchReturnNextLarger)
                    {
                        return ix < searchItems.Count - 1 ? searchItems[ix + 1].OriginalIndex : searchItems[ix].OriginalIndex;
                    }
                }
            }
            return -1;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            Stopwatch sw = null;
            if (context.Debug)
            {
                sw = new Stopwatch();
                sw.Start();
            }
            ValidateArguments(arguments, 3);
            var lookupValue = arguments.ElementAt(0).Value;
            
            // lookup range
            if (!arguments.ElementAt(1).IsExcelRange) return CreateResult(eErrorType.Value);
            var lookupRange = arguments.ElementAt(1).ValueAsRangeInfo;
            var lookupDirection = GetLookupDirection(lookupRange);
            if(lookupRange.Size.NumberOfRows > 1 && lookupRange.Size.NumberOfCols > 1) return CreateResult(eErrorType.Value);

            // return range
            if (!arguments.ElementAt(2).IsExcelRange) return CreateResult(eErrorType.Value);
            var returnArray = arguments.ElementAt(2).ValueAsRangeInfo;
            var notFoundText = string.Empty;

            // not found text
            if (arguments.Count() > 3 && arguments.ElementAt(3) != null)
            {
                notFoundText = ArgToString(arguments, 3);
            }

            // match mode
            var matchMode = XLookupMatchMode.ExactMatch;
            if (arguments.Count() > 4 && arguments.ElementAt(4) != null)
            {
                var mm = ArgToInt(arguments, 4);
                matchMode = GetMatchMode(mm);
            }

            // search mode
            var searchMode = XLookupSearchMode.StartingAtFirst;
            if (arguments.Count() > 5)
            {
                var sm = ArgToInt(arguments, 5);
                searchMode = GetSearchMode(sm);
            }
            int ix;
            if(searchMode == XLookupSearchMode.BinarySearchAscending || searchMode == XLookupSearchMode.BinarySearchDescending)
            {
                ix = GetMatchIndex(lookupValue, lookupRange, searchMode, matchMode);
                if(ix < 0)
                {
                    var ix1 = ~ix;
                    if (matchMode == XLookupMatchMode.ExactMatchReturnNextSmaller)
                    { 
                        ix = ix1 - 1;
                    }
                    else if(matchMode == XLookupMatchMode.ExactMatchReturnNextLarger)
                    {
                        var adjustment = (searchMode == XLookupSearchMode.BinarySearchDescending) ? -1 : 1;
                        var max = returnArray.Size.NumberOfRows > returnArray.Size.NumberOfCols ?
                            returnArray.Size.NumberOfRows : returnArray.Size.NumberOfCols;
                        ix = ix1 >= max ? ix1 : ix1 + adjustment;
                    }
                    
                }
            }
            else
            {
                var lookupArray = GetLookupArray(lookupRange, lookupDirection);
                var sortedLookupArray = GetSortedArray(lookupArray);
                ix = GetMatchIndex(lookupValue, sortedLookupArray, searchMode, matchMode);
            }
            
            if (ix < 0 || ix > ((lookupDirection == XlookupRangeDirection.Vertical) ? returnArray.Size.NumberOfRows - 1 : returnArray.Size.NumberOfCols - 1))
            {
                return string.IsNullOrEmpty(notFoundText) ? CreateResult(eErrorType.NA) : CreateResult(notFoundText, DataType.String);
            }
            var result = default(IRangeInfo);
            if(lookupDirection == XlookupRangeDirection.Vertical)
            {
                var nCols = returnArray.Size.NumberOfCols;
                result = returnArray.GetOffset(ix, 0, ix, nCols - 1);
            }
            else
            {
                var nRows = returnArray.Size.NumberOfRows;
                result = returnArray.GetOffset(0, ix, nRows - 1, ix);
            }
            if(result == null)
            {
                if(string.IsNullOrEmpty(notFoundText))
                {
                    return CreateResult(eErrorType.NA);
                }
                else
                {
                    return CreateResult(notFoundText, DataType.String);
                }
            }
            if (context.Debug)
            {
                sw.Stop();
                context.Configuration.Logger.LogFunction("XLOOKUP", sw.ElapsedMilliseconds);
            }
            return CreateResult(result, DataType.ExcelRange);
        }
    }
}
