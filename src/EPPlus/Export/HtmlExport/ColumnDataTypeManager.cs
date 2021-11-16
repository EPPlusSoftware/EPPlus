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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal static class ColumnDataTypeManager
    {
        private static CompileResultFactory _compileResultFactory = new CompileResultFactory();

        public static class HtmlDataTypes
        {
            public const string Number = "number";
            public const string String = "string";
            public const string Boolean = "boolean";
            public const string DateTime = "datetime";
            public const string TimeOfDay = "timeofday";
        }

        public static string GetColumnDataType(ExcelWorksheet sheet, ExcelRangeBase range, int startRow, int column)
        {
            var rowIndex = startRow;
            var dataType = DataType.Empty;
            while(rowIndex <= range.End.Row)
            {
                var val = sheet.Cells[startRow, column].Value;
                var compileResult = _compileResultFactory.Create(val);
                if (compileResult.DataType != DataType.Empty) return GetHtmlDataType(compileResult.DataType);
                rowIndex++;
            }
            return GetHtmlDataType(dataType);
        }

        private static string GetHtmlDataType(DataType dataType)
        {
            switch(dataType)
            {
                case DataType.Integer:
                case DataType.Decimal:
                    return HtmlDataTypes.Number;
                case DataType.String:
                    return HtmlDataTypes.String;
                case DataType.Boolean:
                    return HtmlDataTypes.Boolean;
                case DataType.Time:
                    return HtmlDataTypes.TimeOfDay;
                case DataType.Date:
                    return HtmlDataTypes.DateTime;
                default:
                    return HtmlDataTypes.String;
            }
        }
    }
}
