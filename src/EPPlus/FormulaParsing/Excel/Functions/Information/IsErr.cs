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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
    internal class IsErr : ErrorHandlingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var isError = new IsError();
            var result = isError.Execute(arguments, context);
            if ((bool) result.Result)
            {
                var arg = GetFirstValue(arguments);
                if (arg is ExcelDataProvider.IRangeInfo)
                {
                    var r = (ExcelDataProvider.IRangeInfo)arg;
                    var e=r.GetValue(r.Address._fromRow, r.Address._fromCol) as ExcelErrorValue;
                    if (e !=null && e.Type==eErrorType.NA)
                    {
                        return CreateResult(false, DataType.Boolean);
                    }
                }
                else
                {
                    if (arg is ExcelErrorValue && ((ExcelErrorValue)arg).Type==eErrorType.NA)
                    {
                        return CreateResult(false, DataType.Boolean);
                    }
                }
            }
            return result;
        }

        public override CompileResult HandleError(string errorCode)
        {
            return CreateResult(true, DataType.Boolean);
        }
    }
}
