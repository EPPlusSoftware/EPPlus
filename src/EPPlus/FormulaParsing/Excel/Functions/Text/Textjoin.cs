/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/18/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Text,
        EPPlusVersion = "5.2",
        IntroducedInExcelVersion = "2019",
        Description = "Joins together two or more text strings, separated by a delimiter")]
    internal class Textjoin : ExcelFunction
    {

        private readonly int MaxReturnLength = 32767;
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 3;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var delimiter = ArgToString(arguments, 0);
            var ignoreEmpty = ArgToBool(arguments, 1);
            var str = new StringBuilder();
            for(var x = 2; x < arguments.Count && x < 252; x++)
            {
                var arg = arguments[x];
                if(arg.IsExcelRange)
                {
                    foreach(var cell in arg.ValueAsRangeInfo)
                    {
                        if(cell.Value != null && cell.Value is ExcelErrorValue eev)
                        {
                            return CompileResult.GetErrorResult(eev.Type);
                        }
                        var val = cell.Value != null ? cell.Value.ToString() : string.Empty;
                        if (ignoreEmpty && string.IsNullOrEmpty(val)) continue;
                        str.Append(val);
                        str.Append(delimiter);
                        if (str.Length > MaxReturnLength) return CompileResult.GetErrorResult(eErrorType.Value);
                    }
                }
                else if(arg.Value is IEnumerable<FunctionArgument>)
                {
                    var items = arg.Value as IEnumerable<FunctionArgument>;
                    if(items != null)
                    {
                        foreach(var item in items)
                        {
                            if (item.Value != null && item.Value is ExcelErrorValue eev)
                            {
                                return CompileResult.GetErrorResult(eev.Type);
                            }
                            var val = item.Value != null ? item.Value.ToString() : string.Empty;
                            if (ignoreEmpty && string.IsNullOrEmpty(val)) continue;
                            str.Append(val);
                            str.Append(delimiter);
                            if (str.Length > MaxReturnLength) return CompileResult.GetErrorResult(eErrorType.Value);
                        }
                    }
                }
                else
                {
                    if (arg.Value != null && arg.Value is ExcelErrorValue eev)
                    {
                        return CompileResult.GetErrorResult(eev.Type);
                    }
                    var val = arg.Value != null ? arg.Value.ToString() : string.Empty;
                    if (ignoreEmpty && string.IsNullOrEmpty(val)) continue;
                    str.Append(val);
                    str.Append(delimiter);
                    if (str.Length > MaxReturnLength) return CompileResult.GetErrorResult(eErrorType.Value);
                }
            }
            var resultString = str.ToString().TrimEnd(delimiter.ToCharArray());
            return CreateResult(resultString, DataType.String);
        }
    }
}
