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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Logical,
        EPPlusVersion = "4",
        Description = "Returns the logical value FALSE")]
    internal class Or : ExcelFunction
    {
        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var argsChecked = 0;
            for (var x = 0; x < arguments.Count(); x++)
            {
                var arg = arguments.ElementAt(x);
                if (arg.IsExcelRange)
                {
                    var range = arg.ValueAsRangeInfo;
                    foreach (var cell in range)
                    {
                        var v = cell.Value;
                        if (v is ExcelErrorValue)
                        {
                            return CreateResult(v, DataType.ExcelError);
                        }
                        else if (!(v is string))
                        {
                            var bVal = ConvertUtil.GetValueDouble(v);
                            argsChecked++;
                            if (bVal != 0d) return CreateResult(true, DataType.Boolean);
                        }
                    }
                }
                else
                {
                    var singleArg = arguments.ElementAt(x);
                    var v = singleArg.Value;
                    if (v == null)
                    {
                        continue;
                    }
                    else if (v is ExcelErrorValue)
                    {
                        return CreateResult(v, DataType.ExcelError);
                    }
                    else if (singleArg.Address == null && v is string)
                    {
                        if (string.IsNullOrEmpty(v.ToString())) continue;
                        if (bool.TryParse(v.ToString(), out bool res))
                        {
                            argsChecked++;
                            if (res) return CreateResult(true, DataType.Boolean);
                        }
                    }
                    else if (!(v is string))
                    {
                        var bVal = ConvertUtil.GetValueDouble(v);
                        argsChecked++;
                        if (bVal != 0d) return CreateResult(true, DataType.Boolean);
                    }
                }
            }
            if (argsChecked == 0)
            {
                return CreateResult(eErrorType.Value);
            }
            return new CompileResult(false, DataType.Boolean);
        }
    }
}
