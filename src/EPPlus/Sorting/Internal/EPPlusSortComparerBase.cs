/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/07/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Sorting.Internal
{
    internal abstract class EPPlusSortComparerBase<T1, T2> : IComparer<T1>
        where T1 : SortItemBase<T2>
    { 


        public EPPlusSortComparerBase(bool[] descending, Dictionary<int, string[]> customLists, CultureInfo culture = null, CompareOptions compareOptions = CompareOptions.None)
        {
            Descending = descending;
            CustomLists = customLists;
            Culture = culture;
            CompOptions = compareOptions;
        }

        protected const int CustomListNotFound = -1;

        protected bool[] Descending { get; private set; }
        protected Dictionary<int, string[]> CustomLists { get; private set; }

        protected CultureInfo Culture { get; private set; }

        protected CompareOptions CompOptions { get; private set; }

        protected int GetSortWeightByCustomList(string val, string[] list)
        {
            if (list == null || list.Count() == 0) return -1;
            var ignoreCase = CompOptions == CompareOptions.IgnoreCase || CompOptions == CompareOptions.OrdinalIgnoreCase;
            for (var x = 0; x < list.Length; x++)
            {
                if (string.Compare(val, list[x], ignoreCase, Culture) == 0)
                    return x;
            }
            return CustomListNotFound;
        }

        protected int CompareObjects(object x1, object y1)
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

        public abstract int Compare(T1 x, T1 y);
    }
}
