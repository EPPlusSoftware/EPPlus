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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "7.0",
        Description = "Returns the one-tailed P-value of a z-test.")]
    internal class Ztest : NormalDistributionBase
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments.Count > 3)
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            else
            {

                var r1 = arguments[0].ValueAsRangeInfo;
                var numbers = RangeFlattener.FlattenRange(r1, false);
                var value = ArgToDecimal(arguments[1].Value, out ExcelErrorValue e1);
                if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
                var stdev = new Stdev().StandardDeviation(numbers.Select(x => x.Value));
                var result = 0d;
                if (stdev.Result is ExcelErrorValue)
                {
                    return stdev;
                }
                double numbersMean = numbers.AverageKahan();
                var z = (numbersMean - value) / (stdev.ResultNumeric / Math.Sqrt(numbers.Count));
                if (arguments.Count < 3)
                {
                    result = 1 - CumulativeDistribution(z, 0, 1);
                }
                else
                {
                    var sigma = ArgToDecimal(arguments[2].Value, out ExcelErrorValue e2);
                    if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
                    if (sigma <= 0)
                    {
                        return CompileResult.GetErrorResult(eErrorType.Num);
                    }
                    else
                    {
                        z = (numbersMean - value) / (sigma / Math.Sqrt(numbers.Count));
                        result = 1 - CumulativeDistribution(z, 0, 1);
                    }
                }
                return CreateResult(result, DataType.Decimal);
            }
        }
    }
}
