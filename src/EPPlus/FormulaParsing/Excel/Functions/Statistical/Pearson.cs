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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "6.0",
        Description = "Returns the Pearson product moment correlation coefficient.")]
    internal class Pearson : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var arg1 = arguments.ElementAt(0);
            var arg2 = arguments.ElementAt(1);
            var array1 = ArgsToDoubleEnumerable(new FunctionArgument[] { arg1 }, context).Select(x => x.Value).ToArray();
            var array2 = ArgsToDoubleEnumerable(new FunctionArgument[] { arg2 }, context).Select(x => x.Value).ToArray();
            if (array1.Count() != array2.Count()) return CreateResult(eErrorType.NA);
            if (!array1.Any()) return CreateResult(eErrorType.NA);
            var result = PearsonImpl(array1, array2);
            return CreateResult(result, DataType.Decimal);
        }

        internal static double PearsonImpl(IEnumerable<double> arr1, IEnumerable<double> arr2)
        {
            var avg1 = arr1.Average();
            var avg2 = arr2.Average();
            var length = arr1.Count();
            var number = 0d;
            double d1 = 0d, d2 = 0d;
            for(var x = 0; x < length; x++)
            {
                number += (arr1.ElementAt(x) - avg1) * (arr2.ElementAt(x) - avg2);
                d1 += System.Math.Pow(arr1.ElementAt(x) - avg1, 2);
                d2 += System.Math.Pow(arr2.ElementAt(x) - avg2, 2);
            }
            return number / System.Math.Sqrt(d1 * d2);
        }
    }
}
