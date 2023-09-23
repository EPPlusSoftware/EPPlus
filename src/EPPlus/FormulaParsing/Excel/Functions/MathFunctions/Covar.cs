/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/12/2020         EPPlus Software AB       Version 5.5
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "5.5",
        Description = "Returns covariance, the average of the products of deviations for each data point pair in two data sets.")]
    internal class Covar : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var array1 = ArgsToDoubleEnumerable(arguments.Take(1), context, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var array2 = ArgsToDoubleEnumerable(arguments.Skip(1).Take(1), context, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            if (array1.Count != array2.Count) return CreateResult(eErrorType.NA);
            if (array1.Count == 0) return CreateResult(eErrorType.Div0);
            var avg1 = array1.AverageKahan();
            var avg2 = array2.AverageKahan();
            var result = 0d;
            for(var x = 0; x < array1.Count; x++)
            {
                result += (array1[x] - avg1) * (array2[x] - avg2);
            }
            result /= array1.Count;
            return CreateResult(result, DataType.Decimal);
        }
    }
}
