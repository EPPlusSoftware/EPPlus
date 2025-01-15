/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  14/1/2025         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/
using System;

namespace OfficeOpenXml.Utils
{
    internal static class ComparerUtil
    {
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
