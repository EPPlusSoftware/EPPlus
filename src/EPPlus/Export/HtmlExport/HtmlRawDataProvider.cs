/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/07/2021         EPPlus Software AB       Added Html Export
 *************************************************************************************************/
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal static class HtmlRawDataProvider
    {
        private static readonly DateTime JsBaseDate = new DateTime(1970, 1, 1);
        internal static string GetHtmlDataTypeFromValue(object value)
        {
            var t = value.GetType();
            var tc = Type.GetTypeCode(t);
            switch (tc)
            {
                case TypeCode.String:
                    return ColumnDataTypeManager.HtmlDataTypes.String;
                case TypeCode.Boolean:
                    return ColumnDataTypeManager.HtmlDataTypes.Boolean;
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Single:
                    return ColumnDataTypeManager.HtmlDataTypes.Number;
                case TypeCode.DateTime:
                    return ColumnDataTypeManager.HtmlDataTypes.DateTime;
                default:
                    if(value is TimeSpan)
                    {
                        return ColumnDataTypeManager.HtmlDataTypes.TimeSpan;
                    }
                    return ColumnDataTypeManager.HtmlDataTypes.String;
            }
        }
        internal static string GetRawValue(object value)
        {
            var t = value.GetType();
            var tc = Type.GetTypeCode(t);
            if (tc == TypeCode.Empty)
            {
                return string.Empty;
            }
            else
            {
                var type = GetHtmlDataTypeFromValue(value);
                return GetRawValue(value, type);
            }
        }
        internal static string GetRawValue(object value, string jsDataType)
        {
            switch(jsDataType)
            {
                case ColumnDataTypeManager.HtmlDataTypes.Boolean:
                    return (ConvertUtil.GetTypedCellValue<bool?>(value, true)??false) ? "1" : "0";
                case ColumnDataTypeManager.HtmlDataTypes.Number:
                    var v = ConvertUtil.GetTypedCellValue<double?>(value, true)?.ToString(CultureInfo.InvariantCulture);
                    return v;
                case ColumnDataTypeManager.HtmlDataTypes.TimeSpan:
                    return ((TimeSpan)value).TotalMilliseconds.ToString(CultureInfo.InvariantCulture);
                case ColumnDataTypeManager.HtmlDataTypes.DateTime:
                    var dt = ConvertUtil.GetTypedCellValue<DateTime?>(value, true);
                    if(dt != null && dt.HasValue)
                    {
                        return dt.Value.Subtract(JsBaseDate).TotalMilliseconds.ToString(CultureInfo.InvariantCulture);
                    }
                    return string.Empty;
                default:
                    return ConvertUtil.GetTypedCellValue<string>(value);

            }
        }
    }
}
