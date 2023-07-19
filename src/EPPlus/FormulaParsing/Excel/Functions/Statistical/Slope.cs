/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  21/06/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
    Category = ExcelFunctionCategory.Statistical,
    EPPlusVersion = "7.0",
    Description = "Returns the slope of the linear regression line through data points in known_y's and known_x's. The slope is the vertical distance divided by the horizontal distance between any two points on the line, which is the rate of change along the regression line.")]

    internal class Slope : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (!arguments[0].IsExcelRange || !arguments[1].IsExcelRange) return CompileResult.GetErrorResult(eErrorType.Value);

            var r1 = arguments[0].ValueAsRangeInfo;
            var r2 = arguments[1].ValueAsRangeInfo;

            if (r1.GetNCells() != r2.GetNCells())
            {
                return CompileResult.GetErrorResult(eErrorType.NA);
            }

            RangeFlattener.GetNumericPairLists(r1, r2, true, out List<double> yValuesFinal, out List<double> xValuesFinal);

            double yMean = yValuesFinal.Select(y => (double)y).Average();
            double xMean = xValuesFinal.Select(x => (double)x).Average();

            var nominator = 0d;
            var denominator = 0d;
            for (var i = 0; i < yValuesFinal.Count; i++)
            {
                var x = xValuesFinal[i];
                var y = yValuesFinal[i];

                denominator += Math.Pow(x - xMean, 2);
                nominator += (x - xMean) * (y - yMean);
            }
            var result = nominator/denominator;
            return CreateResult(result, DataType.Decimal);
        }

    }
}
