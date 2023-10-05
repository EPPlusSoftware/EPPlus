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
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "6.0",
        Description = "Calculates the point at which a line will intersect the y-axis by using existing x-values and y-values.")]
    internal class Intercept : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var knownYs = ArgsToDoubleEnumerable(arguments[0], context, out ExcelErrorValue e1).ToArray();
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var knownXs = ArgsToDoubleEnumerable(arguments[1], context, out ExcelErrorValue e2).ToArray();
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            if (knownYs.Length != knownXs.Length) return CreateResult(eErrorType.NA);
            if (!knownYs.Any()) return CreateResult(eErrorType.NA);
            var result = InterceptImpl(0, knownYs, knownXs);
            return CreateResult(result, DataType.Decimal);
        }

        internal static double InterceptImpl(double x, double[] arrayY, double[] arrayX)
        {
            var avgY = arrayY.AverageKahan();
            var avgX = arrayX.AverageKahan();
            var nItems = arrayY.Length;
            var upperEquationPart = 0d;
            var lowerEquationPart = 0d;
            for (var ix = 0; ix < nItems; ix++)
            {
                upperEquationPart += (arrayX[ix] - avgX) * (arrayY[ix] - avgY);
                lowerEquationPart += System.Math.Pow(arrayX[ix] - avgX, 2);
            }
            return avgY - (upperEquationPart / lowerEquationPart) * avgX;
        }
    }
}
