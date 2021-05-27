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

        public abstract int Compare(T1 x, T1 y);
    }
}
