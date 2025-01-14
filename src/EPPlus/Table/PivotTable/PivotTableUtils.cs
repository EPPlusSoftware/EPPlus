
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
