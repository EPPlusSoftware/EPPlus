/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/7/2021         EPPlus Software AB       EPPlus 5.6
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
    internal class EPPlusSortComparer : IComparer<SortItem<ExcelValue>>
    {
        public EPPlusSortComparer(int[] columns, bool[] descending, Dictionary<int, string[]> customLists, CultureInfo culture = null, CompareOptions compareOptions = CompareOptions.None)
        {
            _columns = columns;
            _descending = descending;
            _customLists = customLists;
            if(culture == null)
            {
                _cultureInfo = CultureInfo.CurrentCulture;
            }
            else
            {
                _cultureInfo = culture;
            }
            _compareOptions = compareOptions;
        }

        private readonly int[] _columns;
        private readonly bool[] _descending;
        private readonly Dictionary<int, string[]> _customLists;
        private readonly CultureInfo _cultureInfo;
        private readonly CompareOptions _compareOptions;
        private const int CustomListNotFound = -1;


        private int GetSortWeightByCustomList(string val, string[] list)
        {
            if (list == null || list.Count() == 0) return -1;
            var ignoreCase = _compareOptions == CompareOptions.IgnoreCase || _compareOptions == CompareOptions.OrdinalIgnoreCase;
            for(var x = 0; x < list.Length; x++)
            {
                if (string.Compare(val, list[x], ignoreCase, _cultureInfo) == 0)
                    return x;
            }
            return CustomListNotFound;
        }
        public int Compare(SortItem<ExcelValue> x, SortItem<ExcelValue> y)
        {
            var ret = 0;
            for (int i = 0; i < _columns.Length; i++)
            {
                var x1 = x.Items[_columns[i]]._value;
                var y1 = y.Items[_columns[i]]._value;
                if(_customLists != null && _customLists.ContainsKey(_columns[i]))
                {
                    var weight1 = GetSortWeightByCustomList(x1.ToString(), _customLists[_columns[i]]);
                    var weight2 = GetSortWeightByCustomList(y1.ToString(), _customLists[_columns[i]]);
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
                
                if (ret != 0) return ret * (_descending[i] ? -1 : 1);
            }
            return 0;
        }
    }
}
