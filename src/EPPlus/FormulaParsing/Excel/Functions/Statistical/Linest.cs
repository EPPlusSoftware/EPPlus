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
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
       Category = ExcelFunctionCategory.Statistical,
       EPPlusVersion = "7.0",
       Description = "The LINEST function calculates the statistics for a line by using the \"least squares\" method to calculate a straight line that best fits your data, and then returns an array that describes the line. You can also combine LINEST with other functions to calculate the statistics for other types of models that are linear in the unknown parameters, including polynomial, logarithmic, exponential, and power series. Because this function returns an array of values, it must be entered as an array formula. Instructions follow the examples in this article.")]
    internal class Linest : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var knownYs = ArgsToDoubleEnumerable(new List<FunctionArgument> { arguments[0] }, context).Select(x => (double)x).ToList();
            List<double> knownXs = default;
            if (arguments.Count == 1)
            {
                knownXs = GetDefaultKnownXs(knownYs.Count);
            }
            else
            {
                knownXs = ArgsToDoubleEnumerable(new List<FunctionArgument> { arguments[1] }, context).Select(x => (double)x).ToList();
            }
            if (knownYs.Count != knownXs.Count)
            {
                return CompileResult.GetErrorResult(eErrorType.Ref);
            }
            var constVar = true;
            if (arguments.Count > 2)
            {
                constVar = ArgToBool(arguments, 2);
            }
            var stats = false;
            if (arguments.Count > 3)
            {
                stats = ArgToBool(arguments, 3);
            }

            var resultRange = CalculateResult(knownYs, knownXs, constVar, stats);
            return CreateResult(resultRange, DataType.ExcelRange);

        }
        private List<double> GetDefaultKnownXs(int count)
        {
            List<double> result = new List<double>();
            for (int i = 1; i <= count; i++)
            {
                result.Add(i);
            }
            return result;
        }
        private InMemoryRange CalculateResult(List <double> knownYs, List <double> knownXs, bool constVar, bool stats)
        {

            var averageY = knownYs.Average();
            var averageX = knownXs.Average();
            
            double nominator = 0d;
            double denominator = 0d;
            for (var i = 0; i < knownYs.Count; i++)
            {
                var y = knownYs[i];
                var x = knownXs[i];

                if (constVar)
                {
                    nominator += (x - averageX) * (y - averageY);
                    denominator += (x - averageX) * (x - averageX);
                }
                else
                {
                    nominator += x * y;
                    denominator += Math.Pow(x, 2);
                }

            }

            var m = nominator / denominator;
            var b = (constVar) ? averageY - (m * averageY) : 0d;

            var resultRange = new InMemoryRange(1, 2);
            resultRange.SetValue(0, 0, m);
            resultRange.SetValue(0, 1, b);
            return resultRange;

        }
    }
}
