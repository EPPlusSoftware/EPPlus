/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/11/2024         EPPlus Software AB           EPPlus v8
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading;
//using Microsoft.Extensions.Primitives;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Text,
        EPPlusVersion = "8",
        Description = "Returns an array of text values from any specified range. It passes text values unchanged, and converts non-text values to text")]
    internal class ArrayToText : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        private const int ConciceFormat = 0;
        private const int StrictFormat = 1;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (!arguments.First().IsExcelRange)
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            var range = arguments.First().ValueAsRangeInfo;
            var format = ConciceFormat;
            if (arguments.Count > 1)
            {
                format = ArgToInt(arguments, 1, out ExcelErrorValue e1);
                if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
                if (format < 0 || format > 1)
                {
                    return CompileResult.GetErrorResult(eErrorType.Value);
                }
            }
            var result = new StringBuilder();
            if (format == StrictFormat)
            {
                result.Append('{');
            }
            var separator = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;

            for (var row = 0; row < range.Size.NumberOfRows; row++)
            {
                for (var col = 0; col < range.Size.NumberOfCols; col++)
                {
                    var val = range.GetOffset(row, col);
                    string strVal = GetStringVal(val, format); 
                    var rowDelimiter = separator.Equals(",") ? ";" : ";"; // Dessa är samma.
                    var colDelimiter = separator.Equals(",") ? "\\" : ",";

                    if (format == ConciceFormat)
                    {
                        if(separator != ",") { rowDelimiter = colDelimiter; } 
                        result.Append(strVal);
                        if (row == range.Size.NumberOfRows - 1 && col == range.Size.NumberOfCols - 1) continue;
                        result.Append(rowDelimiter + " ");                                                 
                    }
                    else
                    {
                        // Strict format
                        result.Append(strVal);
                        if (row == range.Size.NumberOfRows - 1 && col == range.Size.NumberOfCols - 1) continue;

                        if (col < range.Size.NumberOfCols - 1)
                        {
                            result.Append(colDelimiter);
                        }
                        else
                        {
                            result.Append(rowDelimiter);
                        }
                    }
                }
            }
            //var resultStr = format == StrictFormat ? result.ToString().TrimEnd(';', ' ') : result.ToString().TrimEnd(',', ' ');
            var resultStr = result.ToString().TrimEnd(' ');
            if (format == StrictFormat)
            {
                resultStr += "}";
            }
            return CreateResult(resultStr, DataType.String);
        }

        private static string GetStringVal(object val, int format)
        {
            string strVal = string.Empty;
            if (val is bool bVal)
            {
                strVal = bVal ? "TRUE" : "FALSE";
            }
            else if (format == ConciceFormat && val is not null)
            {
                {
                    strVal = val.ToString();
                }
            }
            else if (format == StrictFormat && val is not null)
            {
                if (val is string str)
                {
                    var escaped = str.Replace("\"", "\"\"");
                    strVal = $"\"{escaped}\"";
                }
                else
                {
                    strVal = val.ToString();
                }
            }
            return strVal;
        }
    }
}
