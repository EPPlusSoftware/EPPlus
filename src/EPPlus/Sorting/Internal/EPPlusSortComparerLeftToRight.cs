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
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Sorting.Internal
{
    internal class EPPlusSortComparerLeftToRight : EPPlusSortComparerBase<SortItemLeftToRight<ExcelValue>, ExcelValue>
    {
        public EPPlusSortComparerLeftToRight(int[] rows, bool[] descending, Dictionary<int, string[]> customLists, CultureInfo culture = null, CompareOptions compareOptions = CompareOptions.None)
            : base(descending, customLists, culture, compareOptions)
        {
            _rows = rows;
        }

        private readonly int[] _rows;

        public override int Compare(SortItemLeftToRight<ExcelValue> x, SortItemLeftToRight<ExcelValue> y)
        {
            var ret = 0;
            for (int i = 0; i < _rows.Length; i++)
            {
                var x1 = x.Items[_rows[i]]._value;
                var y1 = y.Items[_rows[i]]._value;
                if (x1 == null && y1 != null) return 1;
                if (x1 != null && y1 == null) return -1;
                if (CustomLists != null && CustomLists.ContainsKey(_rows[i]))
                {
                    var weight1 = GetSortWeightByCustomList(x1.ToString(), CustomLists[_rows[i]]);
                    var weight2 = GetSortWeightByCustomList(y1.ToString(), CustomLists[_rows[i]]);
                    ret = weight1.CompareTo(weight2);
                }
                else
                {
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
                }

                if (ret != 0) return ret * (Descending[i] ? -1 : 1);
            }
            return 0;
        }
    }
}
