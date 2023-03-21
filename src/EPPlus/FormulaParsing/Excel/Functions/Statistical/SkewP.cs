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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
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
        Description = "Calculates the skewness of a distribution based on a population")]
    internal class SkewP : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var numbers = ArgsToDoubleEnumerable(arguments, context).Select(x => x.Value).ToArray();
            var n = numbers.Length;
            var avg = numbers.Average();
            var sd = CalcStandardDev(numbers, avg);
            var i = 0d;
            foreach(var number in numbers)
            {
                i += System.Math.Pow((number - avg)/sd, 3);
            }
            var result = i / n;
            return CreateResult(result, DataType.Decimal);
        }

        private double CalcStandardDev(IEnumerable<double> numbers, double avg)
        {
            var stdDev = 0d;
            foreach (var n in numbers)
            {
                stdDev += System.Math.Pow(n - avg, 2);
            }
            return System.Math.Sqrt(stdDev / numbers.Count());
        }
    }
}
