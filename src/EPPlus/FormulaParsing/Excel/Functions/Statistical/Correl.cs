/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/10/2022         EPPlus Software AB           EPPlus v6
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using MathObj = System.Math;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
            Category = ExcelFunctionCategory.Statistical,
            EPPlusVersion = "6.0",
            Description = "Returns the correlation coefficient of two cell ranges")]
    internal class Correl : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var arg1 = arguments[0];
            var arg2 = arguments[1];
            var list1 = ArgsToDoubleEnumerable(arg1, context, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var list2 = ArgsToDoubleEnumerable(arg2, context, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);

            if (list2.Count != list1.Count) return CreateResult(eErrorType.NA);
            if (list1.SumKahan() == 0 || list2.SumKahan() == 0) return CreateResult(eErrorType.Div0);
            var result = Covar(list1, list2) / StandardDeviation(list1) / StandardDeviation(list2);
            return CreateResult(result, DataType.Decimal);
        }

        private static double StandardDeviation(IList<double> values)
        {
            double avg = values.AverageKahan();
            return MathObj.Sqrt(values.AverageKahan(v => MathObj.Pow(v - avg, 2)));
        }

        private static double Covar(IList<double> array1, IList<double> array2)
        {
            var avg1 = array1.AverageKahan();
            var avg2 = array2.AverageKahan();
            var result = 0d;
            for (var x = 0; x < array1.Count; x++)
            {
                result += (array1[x] - avg1) * (array2[x] - avg2);
            }
            result /= array1.Count;
            return result;
        }
    }
}
