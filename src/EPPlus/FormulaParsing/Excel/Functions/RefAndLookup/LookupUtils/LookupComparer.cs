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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils
{
    internal class LookupComparer : IComparer<object>
    {
        public LookupComparer(LookupMatchMode matchMode)
        {
            _matchMode = matchMode;
        }

        private readonly LookupMatchMode _matchMode;
        private readonly int _sortOrder;
        private readonly ValueMatcher _vm = new WildCardValueMatcher();

        public virtual int Compare(object x, object y, int sortOrder)
        {
            int v = 0;
            if(x == null && y == null)
            {
                return 1;
            }
            if (_matchMode == LookupMatchMode.Wildcard || _matchMode == LookupMatchMode.ExactMatchWithWildcard)
            {
                v = _vm.IsMatch(x, y);
            }
            else
            {
                v = CompareObjects(x, y);
            }
            return v * (sortOrder > -1 ? 1 : -1);
        }

        public virtual int Compare(object x, object y)
        {
            return Compare(x, y, 1);
        }

        internal static int CompareObjects(object x1, object y1)
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
                ret = string.Compare(s1, s2, StringComparison.CurrentCultureIgnoreCase);
            }
            else
            {
                ret = isNumX ? -1 : 1;
            }

            return ret;
        }
    }
}
