/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable
{
    internal class CacheComparer : IEqualityComparer<object>
    {
        public new bool Equals(object x, object y)
        {
			x = GetCaseInsensitiveValue(x);
            y = GetCaseInsensitiveValue(y);
			return x.Equals(y);           
		}

        private static object GetCaseInsensitiveValue(object x)
        {
            if (x == null || x.Equals(ExcelPivotTable.PivotNullValue)) return ExcelPivotTable.PivotNullValue;

			if (x is string sx)
            {
				return sx.ToLower();
			}
            else if (x is char cx)
            {
                return char.ToLower(cx).ToString();
            }
            else if(x is DateTime)
            {
                return x;
            }
            else if(x is TimeSpan ts)
            {
                return DateTime.FromOADate(0).Add(ts);
            }
            if(ConvertUtil.IsExcelNumeric(x))
            {
                return ConvertUtil.GetValueDouble(x);
            }
            return x.ToString().ToLower();
        }

        public int GetHashCode(object obj)
        {
            return GetCaseInsensitiveValue(obj).GetHashCode();
        }
    }
}