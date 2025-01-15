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
using OfficeOpenXml.Utils;
using System;

namespace OfficeOpenXml.Table.PivotTable
{
    internal static class PivotTableUtils
    {
        internal static object GetCaseInsensitiveValue(object x)
        {
            if (x == null || x.Equals(ExcelPivotTable.PivotNullValue) || x == DBNull.Value) return ExcelPivotTable.PivotNullValue;
            var tc = Type.GetTypeCode(x.GetType());
            switch (tc)
            {
                case TypeCode.String:
                    return x.ToString().ToLower();
                case TypeCode.Char:
                    return ((char)x).ToString().ToLower();
                case TypeCode.DateTime:
                case TypeCode.Boolean:
                    return x;
                case TypeCode.Object:
                    if (x is TimeSpan ts)
                    {
                        return DateTime.FromOADate(0).Add(ts);
                    }
                    return x.ToString().ToLower();
                default:
                    if (ConvertUtil.IsExcelNumeric(x))
                    {
                        return ConvertUtil.GetValueDouble(x);
                    }
                    return x.ToString().ToLower();
            }
        }

    }
}
