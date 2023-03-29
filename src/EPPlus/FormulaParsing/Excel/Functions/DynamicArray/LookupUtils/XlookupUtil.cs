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
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DynamicArray.LookupUtils
{
    internal static class XlookupUtil
    {
        internal static LookupRangeDirection GetLookupDirection(IRangeInfo lookupRange)
        {
            var result = LookupRangeDirection.Vertical;
            if (lookupRange.Size.NumberOfCols > 1)
            {
                result = LookupRangeDirection.Horizontal;
            }
            return result;
        }

        internal static LookupMatchMode GetMatchMode(int mm)
        {
            switch (mm)
            {
                case 0:
                    return LookupMatchMode.ExactMatch;
                case -1:
                    return LookupMatchMode.ExactMatchReturnNextSmaller;
                case 1:
                    return LookupMatchMode.ExactMatchReturnNextLarger;
                case 2:
                    return LookupMatchMode.Wildcard;
                default:
                    throw new ArgumentException("Invalid match mode: " + mm.ToString());
            }
        }


        internal static LookupSearchMode GetSearchMode(int sm)
        {
            switch (sm)
            {
                case 1:
                    return LookupSearchMode.StartingAtFirst;
                case -1:
                    return LookupSearchMode.ReverseStartingAtLast;
                case 2:
                    return LookupSearchMode.BinarySearchAscending;
                case -2:
                    return LookupSearchMode.BinarySearchDescending;
                default:
                    throw new ArgumentException("Invalid search mode: " + sm.ToString());

            }
        }

        internal static List<object> GetLookupArray(IRangeInfo lookupRange, LookupRangeDirection direction)
        {
            var arr = new List<object>();
            if (direction == LookupRangeDirection.Vertical)
            {
                var dimensionRows = lookupRange.Worksheet.Dimension.Rows;
                var maxRows = lookupRange.Size.NumberOfRows > dimensionRows ? dimensionRows : lookupRange.Size.NumberOfRows;
                for (var rowIx = 0; rowIx < maxRows; rowIx++)
                {
                    arr.Add(lookupRange.GetOffset(rowIx, 0));
                }
            }
            else
            {
                var dimensionCols = lookupRange.Worksheet.Dimension.Columns;
                var maxCols = lookupRange.Size.NumberOfCols > dimensionCols ? dimensionCols : lookupRange.Size.NumberOfCols;
                for (var colIx = 0; colIx < maxCols; colIx++)
                {
                    var v = lookupRange.GetOffset(0, colIx);
                    arr.Add(v);
                }
            }
            return arr;
        }

        internal static List<LookupSearchItem> GetSortedArray(List<object> arr)
        {
            var sortedList = new List<LookupSearchItem>();
            if (arr == null || arr.Count == 0) return sortedList;
            for (var ix = 0; ix < arr.Count; ix++)
            {
                sortedList.Add(new LookupSearchItem(arr[ix], ix));
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
                ret = d1 < d2 ? -1 : d1 > d2 ? 1 : 0;
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
    }
}
