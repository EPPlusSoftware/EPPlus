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
using OfficeOpenXml.FormulaParsing.Ranges;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Text,
        EPPlusVersion = "8",
        Description = "Returns text from any specified value. It passes text values unchanged, and converts non-text values to text.")]
    internal class ValueToText : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        private const int ConciceFormat = 0;
        private const int StrictFormat = 1;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
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
            var arg0 = arguments[0];
            if (arg0.IsExcelRange)
            {
                var range = arg0.ValueAsRangeInfo;

                var ret = new InMemoryRange(range.Size);
                
                return CreateDynamicArrayResult(ret, DataType.ExcelRange);
            }
            else
            {
                var stringRes = GetStringVal(arg0, format);
                return CreateResult(stringRes, DataType.String);
            }            
        }

        private static string GetStringVal(object val, int format)
        {
            string strVal = string.Empty;
            if(val is bool bVal)
            {
                strVal = bVal ? "TRUE" : "FALSE";
            }
            else if (format == ConciceFormat && val is not null)
            {
                strVal = val.ToString();
            }
            else if(format == StrictFormat && val is not null) 
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
