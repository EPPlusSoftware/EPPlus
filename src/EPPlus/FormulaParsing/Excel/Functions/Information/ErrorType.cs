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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Information,
        EPPlusVersion = "4",
        Description = "Tests a supplied value and returns an integer relating to the supplied value's error type",
        SupportsArrays = true)]
    internal class ErrorType : ExcelFunction
    {
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var error = arguments.ElementAt(0);
            var isErrorFunc = context.Configuration.FunctionRepository.GetFunction("iserror");
            var isErrorResult = isErrorFunc.Execute(arguments, context);
            if (!(bool) isErrorResult.Result)
            {
                return CompileResult.GetErrorResult(eErrorType.NA);
            }
            var errorType = error.ValueAsExcelErrorValue;
            switch (errorType.Type)
            {
                case eErrorType.Null:
                    return CreateResult(1, DataType.Integer);
                case eErrorType.Div0:
                    return CreateResult(2, DataType.Integer);
                case eErrorType.Value:
                    return CreateResult(3, DataType.Integer);
                case eErrorType.Ref:
                    return CreateResult(4, DataType.Integer);
                case eErrorType.Name:
                    return CreateResult(5, DataType.Integer);
                case eErrorType.Num:
                    return CreateResult(6, DataType.Integer);
                case eErrorType.NA:
                    return CreateResult(7, DataType.Integer);
            }
            return CreateResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);
        }
    }
}
