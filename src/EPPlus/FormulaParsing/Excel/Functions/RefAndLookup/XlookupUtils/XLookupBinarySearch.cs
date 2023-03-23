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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.XlookupUtils
{
    internal static class XLookupBinarySearch
    {
        internal static int Search(object s, XlookupSearchItem[] items, IComparer<object> comparer)
        {
            if (items.Length == 0) return -1;
            int low = 0, high = items.Length - 1, mid;

            while (low <= high)
            {
                mid = (low + high) >> 1;

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

        internal static int SearchDesc(object s, XlookupSearchItem[] items, IComparer<object> comparer)
        {
            if (items.Length == 0) return -1;
            int low = 0, high = items.Length - 1, mid;

            while (high >= low)
            {
                mid = (high + low) >> 1;

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
    }
}
