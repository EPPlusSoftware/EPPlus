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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Returns the Kth SMALLEST value from a list of supplied numbers, for a given value K")]
    internal class Small : HiddenValuesHandlingFunction
    {
        public Small()
        {
            IgnoreHiddenValues = false;
            IgnoreErrors = false;
        }
        //public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        //{
        //    return FunctionParameterInformation.IgnoreErrorInPreExecute;
        //}));
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var args = arguments[0];
            
            
            
            var index = ArgToInt(arguments, 1, IgnoreErrors) - 1;
            var values = ArgsToDoubleEnumerable(args, context, x =>
            {
                x.IgnoreNonNumeric = true;
                x.IgnoreHiddenCells = IgnoreHiddenValues;
                x.IgnoreErrors = IgnoreErrors;
            }, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            if (index < 0 || index >= values.Count()) return CompileResult.GetErrorResult(eErrorType.Num);
            var result = values.OrderBy(x => x).ElementAt(index);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
